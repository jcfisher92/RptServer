/**
 * File: PerfAdderStatus.java
 * Description: No description given in the original report.
 *    This is the rewrite of the report to work with the new report server.
 *    The orginal author was Peggy Richter.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/16/2005
 * Last Update: $Id: PerfAdderStatus.java,v 1.9 2014/01/13 15:54:24 tli Exp $
 *
 * History:
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    05/03/2004 - Removed the setting of the m_DistList variable when sending the report notification.  This variable
 *       will get cleaned up before it can be used by the webservice. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    12/18/2003 - modified the program to handle being run as a web service. - jcf
 *
 *    03/19/2002 - Changed the way the connection to Oracle is established.  The program now uses the
 *       getOracleConn() method to retrieve a new connection object.  The connection pool is no longer
 *       used. Also added the cleanup method.- jcf
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;


public class PerfAdderStatus extends Report
{
   private PreparedStatement m_PerfAdder;
   private PreparedStatement m_Sales;
   private PreparedStatement m_DelinqCnt;
   private PreparedStatement m_ActiveCnt;
   private PreparedStatement m_DelinqLst;
   private PreparedStatement m_CustRep;
   private PreparedStatement m_PerfAdderSum; 

   private int m_Month;
   private boolean m_Recalc;
   private int m_Year;

   /**
    * default constructor
    */
   public PerfAdderStatus()
   {
      super();

      m_Year = 2017;
      m_Month = 12;
   }

   /**
    * Build the output file
    * @return boolean true if the file was built, false if not.
    *
    * @throws FileNotFoundException
    */
   public boolean buildOutputFile() throws FileNotFoundException
   {
      StringBuffer Line = new StringBuffer(1024);
      String fileName = null;
      FileOutputStream OutFile = null;
      ResultSet PerfData = null;
      ResultSet SalesData = null;
      ResultSet ActiveData = null;
      ResultSet DelinqLstData = null;
      ResultSet CustRepData = null;
      String custid;
      int perfid;
      double sales = 0;
      double perfamt = 0;
      double proj_sales = 0;
      double sales_goal = 0;
      double diff = 0;
      int activecnt = 0;
      String under26k = null;
      boolean result = true;
      String delinqlst = null;
      int custrepid = 0;
      String custrepname = "";

      fileName = m_RptProc.getUid() + "pasumm.dat";
      m_FileNames.add(fileName);
      OutFile = new FileOutputStream(m_FilePath + fileName, false);

      //
      // Build the Captions
      Line.append("Customer ID\tName\tStatus\tContact\tPhone\tYTD Sales\tProj Sales\tSales Goal\t");
      Line.append("Shortfall\tDelinq Mo\tActive Mo\tUnder 26K\tFrt%\tBase%\tPA%\tAccrued PA\t");
      Line.append("Est Refund\tSales Rep\tRep ID\r\n");

      try {
         setCurAction( "Calculating" );
         PerfData = m_PerfAdder.executeQuery();

         while ( PerfData.next() && m_Status == RptServer.RUNNING ) {
            custid = PerfData.getString("customer_id");
            setCurAction( "Processing " + custid );

            perfid = PerfData.getInt("perf_id");
            proj_sales = PerfData.getDouble("projected_sales");
            sales_goal = PerfData.getDouble("sales_goal");
            delinqlst = null;

            m_Sales.setLong(1, perfid);
            SalesData = m_Sales.executeQuery();

            if ( SalesData.next() ) {
               sales = SalesData.getDouble("sales");
               perfamt = SalesData.getDouble("perfamt");
            }

            diff = proj_sales - sales_goal;
            if ( diff > 0 )
               diff = 0;

            m_ActiveCnt.setLong(1, perfid);
            ActiveData = m_ActiveCnt.executeQuery();

            if ( ActiveData.next() )
               activecnt = ActiveData.getInt("activecnt");

            under26k = PerfData.getString("under26K");

            m_DelinqLst.setLong(1, perfid);
            DelinqLstData = m_DelinqLst.executeQuery();

            while ( DelinqLstData.next() && m_Status == RptServer.RUNNING ) {
               if (delinqlst == null)
                  delinqlst = "" + DelinqLstData.getInt("sales_month");
               else
                  delinqlst = delinqlst + "," + DelinqLstData.getInt("sales_month");
            }

            if (delinqlst == null)
               delinqlst = " ";

            m_CustRep.setString(1, custid);
            CustRepData = m_CustRep.executeQuery();

            if ( CustRepData.next() ) {
               custrepid = CustRepData.getInt("er_id");
               custrepname = CustRepData.getString("repname");
            }

            //Tmp.setMsgText("Before building output line: \r\n");
            Line.append(custid + "\t");
            Line.append(PerfData.getString("name") + "\t");
            Line.append(PerfData.getString("status") + "\t");
            Line.append(PerfData.getString("contact") + "\t");
            Line.append(PerfData.getString("contactphone") + "\t");
            Line.append(sales + "\t");
            Line.append(proj_sales + "\t");
            Line.append(sales_goal + "\t");
            Line.append(diff + "\t");
            Line.append(delinqlst + "\t");
            Line.append(activecnt + "\t");
            Line.append(under26k + "\t");
            Line.append(PerfData.getFloat("frt") + "\t");
            Line.append(PerfData.getFloat("vol") + "\t");
            Line.append(PerfData.getFloat("perf_adder_pct") + "\t");
            Line.append(perfamt + "\t");
            Line.append(PerfData.getFloat("credit_amt") + "\t");
            Line.append(custrepname + "\t");
            Line.append(custrepid + "\r\n");

            DbUtils.closeDbConn(null, null, SalesData);
            DbUtils.closeDbConn(null, null, ActiveData);
            DbUtils.closeDbConn(null, null, DelinqLstData);
            DbUtils.closeDbConn(null, null, CustRepData);

            OutFile.write(Line.toString().getBytes());
            Line.delete(0, Line.length());
         }

         DbUtils.closeDbConn(null, null, PerfData);

         setCurAction( "Complete");
         result = true;
      }

      catch( Exception ex ) {
         log.error("[PerfAdderStatus]", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n" + ex.getMessage());
         result = false;
      }

      finally  {
         try {
            OutFile.close();
            OutFile = null;
         }

         catch ( Exception ex) {

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
      closeStatements();

      m_PerfAdder = null;
      m_Sales = null;
      m_DelinqCnt = null;
      m_ActiveCnt = null;
      m_DelinqLst = null;
      m_CustRep = null;
      m_PerfAdderSum = null;
   }

   //
   // Close queries
   public void closeStatements()
   {
      DbUtils.closeDbConn(null, m_PerfAdderSum, null);
      
      try {
         if ( m_PerfAdder != null )
            m_PerfAdder.close();
      }
      catch ( Exception ex ) {

      }

      try {
         if ( m_Sales != null )
            m_Sales.close();
      }
      catch ( Exception ex ) {

      }

      try {
         if ( m_DelinqCnt != null )
            m_DelinqCnt.close();
      }
      catch ( Exception ex ) {

      }

      try {
         if ( m_ActiveCnt != null )
            m_ActiveCnt.close();
      }
      catch ( Exception ex ) {

      }

      try {
         if ( m_DelinqLst != null )
            m_DelinqLst.close();
      }
      catch ( Exception ex ) {

      }

      try {
         if ( m_CustRep != null )
            m_CustRep.close();
      }
      catch ( Exception ex ) {

      }
   }

   /**
    * Creates the report.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() ) {
            if ( m_Recalc )
               recalc(m_Month, m_Year);

            created = buildOutputFile();
         }
      }

      catch ( Exception ex ) {
         log.fatal("[PerfAdderStatus]", ex);
      }

      finally {
        closeStatements();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }


   /**
    * Prepare the sql statements
    *
    * @return true if the statements were prepared, false if not.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();

      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("perf_id, perf_adder.customer_id, cust_view.name, status, ");
            sql.append("sales_goal, decode(under_26k, true, 'yes', ' ') as under26k, ");
            sql.append("perf_adder_pct, sales_rep, projected_sales, credit_amt, ");
            sql.append("ejd.cust_procs.adder_pct(cust_view.customer_id, 'FREIGHT') frt, ");
            sql.append("ejd.cust_procs.adder_pct(cust_view.customer_id, 'BASE COST ADDER') vol, ");
            sql.append("ejd.cust_procs.ap_contact_name(cust_view.customer_id) contact, ");
            sql.append("ejd.emery_utils.format_phone(cust_procs.ap_cont_phone(cust_view.customer_id)) contactphone ");
            sql.append("from perf_adder ");
            sql.append("join customer on customer.customer_id = perf_adder.customer_id and customer.cust_status_id in (1,2) ");
            sql.append("join cust_view on cust_view.customer_id =  customer.customer_id ");
            sql.append(String.format("where sales_year = %d ", m_Year));
            sql.append("order by customer.name ");           
            m_PerfAdder = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select nvl(sum(sales),0) sales, nvl(sum(perf_adder_amt),0) perfamt ");
            sql.append("from perf_mo ");
            sql.append("where perf_id = ? and sales_month <= " + m_Month);            
            m_Sales = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select nvl(count(*),0) as delinqcnt ");
            sql.append("from perf_mo ");
            sql.append("where perf_id = ? and credit_ok and sales_month <= " + m_Month);
            m_DelinqCnt = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select count(*) as activecnt ");
            sql.append("from perf_mo ");
            sql.append("where perf_id = ? and sales_month <= " + m_Month + " and is_active ");
            m_ActiveCnt = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select sales_month ");
            sql.append("from perf_mo ");
            sql.append("where perf_id = ? and credit_ok and sales_month <= " + m_Month );
            sql.append("order by sales_month ");            
            m_DelinqLst = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select cust_rep.er_id, (emery_rep.first || ' ' || emery_rep.last) as repname ");
            sql.append("from cust_rep ");
            sql.append("join emery_rep on emery_rep.er_id = cust_rep.er_id ");
            sql.append("join emery_rep_type on emery_rep_type.rep_type_id = cust_rep.rep_type_id and emery_rep_type.description = 'SALES REP' ");
            sql.append("where cust_rep.customer_id = ? ");            
            m_CustRep = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append(" select * from ejd_sa_procs.perf_adder_summary(?,?)");
            m_PerfAdderSum = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;
         }

         catch ( Exception ex ) {
            log.fatal("[PerfAdderStatus]", ex);
         }
      }
      else
         log.fatal("[PerfAdderStatus] null db connection");

      return isPrepared;
   }

   /**
    * If requested, recalculate the month's performance adder results
    *
    * @param mnth
    * @param year
    */
   public void recalc(int month, int year)
   {
      ResultSet rs = null;
      boolean success = true;
      String msg = null;
      
      try {
         log.info("[PerfAdderStatus] Recalculating performance adder results... " + m_RptProc.getUid());
         m_PerfAdderSum.setInt(1, month);
         m_PerfAdderSum.setInt(2, year);
         
         rs = m_PerfAdderSum.executeQuery();
         
         if ( rs.next() ) {
            // should check the return value for success or failure
            success = rs.getInt("code") == 1;
            msg = rs.getString("msg");            
         }
         else {
            success = false;
            msg = "unkown error";
         }
         
         if ( !success ) {
            log.error("[PerfAdderStatus]\nError executing perf adder recalc\n" + msg + '\n' + m_RptProc.getUid());
         }
      }
      
      catch (SQLException sqlex){
          log.error("[PerfAdderStatus]\nError executing perf adder recalc\n" + sqlex.toString() + '\n' + m_RptProc.getUid());
      }      
   }

   /**
    * Sets the parameters for the report
    *    param(0) = sales year
    *    param(1) = sales month
    *    param(2) = boolean
    *
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      m_Year = Integer.parseInt(params.get(0).value);
      m_Month = Integer.parseInt(params.get(1).value);
      m_Recalc = Boolean.parseBoolean(params.get(2).value);
   }
}
