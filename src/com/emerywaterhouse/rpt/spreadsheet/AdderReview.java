/**
 * File: AdderReview.java
 * Description: The adder review report.  Transfered from the mt server to the report server.  The
 *    original author was Peggy Richter.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/09/2005
 * Last Update: $Id: AdderReview.java,v 1.9 2009/02/17 22:45:35 jfisher Exp $
 * 
 * History
 *    $Log: AdderReview.java,v $
 *    Revision 1.9  2009/02/17 22:45:35  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.GregorianCalendar;

import oracle.jdbc.OracleCallableStatement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;


public class AdderReview extends Report
{
   private GregorianCalendar m_EndDate = null;
   private int m_SalesYear = -1;
   private int m_BudgetYear = -1;   
   private boolean m_Spreadsheet = false;

   private double m_CurSales = -1;
   private double m_ProjSales = -1;

   private StringBuffer m_Lines;
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;

   private PreparedStatement m_Budget = null;               //gets the customer's yearly budget
   private PreparedStatement m_Cust = null;                 //return a list of customers to report
   private PreparedStatement m_Eow = null;                  //returns whether a customer ships every other week
   private PreparedStatement m_YtdSales = null;             //returns customer's ytd sales

   private OracleCallableStatement m_AcctSales = null;      //returns the total & projected sales for the account
   private OracleCallableStatement m_ActiveMonths = null;   //returns the number of months the customer has been active during the reporting period
   private OracleCallableStatement m_Adders = null;         //analyzes customer's adders based on sales level
   private OracleCallableStatement m_Under26K = null;       //Does customer have under 26K service charge?
      
   /**
    * default constructor
    */
   public AdderReview()
   {
      super();
   }
   
   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
      super.finalize();
   }

   /**
    * Calculates the current and projected sales at the customer's account level
    * @param custid
    * @param salesyear
    */
   private void acctSales( String custid, int salesyear )
   {
      m_CurSales = -1;
      m_ProjSales = -1;

      try {
         m_AcctSales.setString(1, custid);
         m_AcctSales.setInt(2, salesyear);
         m_AcctSales.execute();
         m_CurSales = m_AcctSales.getDouble(3);
         m_ProjSales = m_AcctSales.getDouble(4);
      }
      catch ( Exception e ) {
         log.error( e );
      }
   }

   /**
    * Returns the number of months the customer was active during the sales reporting period
    * @param custid
    * @return Active Months int
    */
   private int activeMonths( String custid )
   {
      java.sql.Date endDate = new java.sql.Date(m_EndDate.getTimeInMillis());

      try {
         m_ActiveMonths.setString(1, custid);
         m_ActiveMonths.setDate(2, endDate);
         m_ActiveMonths.execute();
         return m_ActiveMonths.getInt(3);
      }
      catch ( Exception e ) {
         log.error( e );
      }

      return -1;
   }

   /**
    * Builds the output file
    * @return boolean.  True if the file was created, false if not.
    * @throws FileNotFoundException
    */
   public boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      boolean result = true;
      ResultSet rs = null;
      String custid = null;
      int activeMo = 12;
      double sales = 0;
      double projsales = 0;   
      double budget = 0;
      short col;
      short rownum = (short)1;
      XSSFRow Row = null;
      
      m_FileNames.add(m_RptProc.getUid() + "adderreview" + (m_Spreadsheet ? ".xlsx": ".dat"));      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();

      try {
         setCurAction( "Querying Customers" );         
         rs = m_Cust.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            custid = rs.getString("customer_id");
            setCurAction( "Processing " + custid + " " + rs.getString("name"));            
            activeMo = activeMonths(custid);
            sales = ytdSales(custid, m_SalesYear);
            
            if ( activeMo == 0 )
            	projsales = 0;
            else
            	projsales = projectedSales(activeMo, sales);
            
            acctSales(custid, m_SalesYear);
            budget = getBudget(custid, m_BudgetYear);

            m_Adders.setString(1, custid);
            m_Adders.setDouble(2, m_ProjSales);
            m_Adders.setDouble(3, projsales);
            m_Adders.execute();

            if ( m_Spreadsheet ) {
               Row = m_Sheet.createRow(rownum++);
               col = (short)0;
               createCell(Row, col++, custid);
               createCell(Row, col++, rs.getString("name"));
               createCell(Row, col++, rs.getString("status"));
               createCell(Row, col++, rs.getString("repname"));
               createCell(Row, col++, rs.getString("division"));
               createCell(Row, col++, m_Adders.getString(9));
               createCell(Row, col++, m_Adders.getInt(10));
               createCell(Row, col++, activeMo);
               createCell(Row, col++, sales);
               createCell(Row, col++, projsales);
               createCell(Row, col++, m_CurSales);
               createCell(Row, col++, m_ProjSales);
               createCell(Row, col++, budget);
               createCell(Row, col++, rs.getString("top100"));
               createCell(Row, col++, getEOW(custid));
               createCell(Row, col++, getUnder26K(custid));
               createCell(Row, col++, rs.getInt("min_order"));
               createCell(Row, col++, m_Adders.getDouble(4));
               createCell(Row, col++, m_Adders.getDouble(5));
               createCell(Row, col++, m_Adders.getDouble(11));
               createCell(Row, col++, m_Adders.getDouble(11) + m_Adders.getDouble(7));

               if ( m_Adders.getDouble(7) != 0 )
                  createCell(Row, col++, m_Adders.getDouble(5));            //recommended change in bc + pa
               else
                  col++;

               createCell(Row, col++, m_Adders.getDouble(6));
               createCell(Row, col++, m_Adders.getDouble(8));
            }
            else {
               m_Lines.append(custid + "\t");
               m_Lines.append(rs.getString("name") + "\t");
               m_Lines.append(rs.getString("status") + "\t");
               m_Lines.append(rs.getString("repname") + "\t");
               m_Lines.append(rs.getString("division") + "\t");
               m_Lines.append(m_Adders.getString(9) + "\t");                   //zip code
               m_Lines.append(m_Adders.getInt(10) + "\t");                      //frt zone
               m_Lines.append(activeMo + "\t");
               m_Lines.append(sales + "\t");
               m_Lines.append(projsales + "\t");
               m_Lines.append(m_CurSales + "\t");
               m_Lines.append(m_ProjSales + "\t");
               m_Lines.append(getBudget(custid, m_BudgetYear) + "\t");
               m_Lines.append(rs.getString("top100") + "\t");
               m_Lines.append(getEOW(custid) + "\t");
               m_Lines.append(getUnder26K(custid) + "\t");
               m_Lines.append(rs.getInt("min_order") + "\t");
               m_Lines.append(m_Adders.getDouble(4) + "\t");            //current base cost adder
               m_Lines.append(m_Adders.getDouble(5) + "\t");            //current performance adder
               m_Lines.append(m_Adders.getDouble(11) + "\t");           //total bc + pa
               m_Lines.append((m_Adders.getDouble(11) + m_Adders.getDouble(7)) + "\t"); //recommended bc + pa

               if ( m_Adders.getDouble(7) != 0 )
                  m_Lines.append(m_Adders.getDouble(7) + "\t");            //recommended change in bc + pa
               else
                  m_Lines.append("\t");

               m_Lines.append(m_Adders.getDouble(6) + "\t");            //current freight adder
               m_Lines.append(m_Adders.getDouble(8) + "\r\n");          //recommended freight
               
               outFile.write(m_Lines.toString().getBytes());
               m_Lines.delete(0, m_Lines.length());
            }
         }

         if ( m_Spreadsheet )
            m_WrkBk.write(outFile);
            
         setCurAction( "Complete");         
         rs.close();
      }

      catch( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         
         result = false;
      }

      finally {
         m_Lines = null;

         try {
            outFile.close();
            outFile = null;
         }
         catch ( Exception e ) {
            log.error( e );
         }         
      }
      
      return result;
   }

   /**
    * Marks objects for garbage collection
    */
   protected void cleanup()
   {
      closeStatement(m_Budget);
      closeStatement(m_Cust);
      closeStatement(m_Eow);
      closeStatement(m_YtdSales);
      closeStatement(m_AcctSales);
      closeStatement(m_ActiveMonths);
      closeStatement(m_Adders);
      closeStatement(m_Under26K);

      m_Budget = null;
      m_Cust = null;
      m_Eow = null;
      m_YtdSales = null;

      m_AcctSales = null;
      m_ActiveMonths = null;
      m_Adders = null;
      m_Under26K = null;

      m_Lines = null;
      m_WrkBk = null;
      m_Sheet = null;
   }

   /**
    * Closes the statement.
    *
    * @param stmt The statement to close.
    */
   private void closeStatement(Statement stmt)
   {
      if ( stmt != null ) {
         try {
            stmt.close();
         }

         catch ( Exception ex ) {
            ;
         }

         finally {
            stmt = null;
         }
      }
   }

   /**
    * Creates the report file.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
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
    * Returns the customers current warehouse budget
    * @param custid The emery customer id
    * @param budgetYear The budget year
    * @return double The budget amount
    */
   private double getBudget( String custid, int budgetYear )
   {
      double budget = 0;
      ResultSet rs = null;

      try {
         m_Budget.setString(1, custid);
         m_Budget.setInt(2, budgetYear);
         rs = m_Budget.executeQuery();

         while ( rs.next() ) {
            budget = rs.getDouble("budget");
         }
      }
      
      catch ( Exception e ) {
         log.error("exception", e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }
      
      return budget;
   }

   /**
    * returns a Y if the customer ships every other week instead of weekly
    * @param custid
    * @return "Y" if every other week
    */
   private String getEOW( String custid )
   {
      String eow = " ";
      ResultSet rs = null;

      try {
         m_Eow.setString(1, custid);
         rs = m_Eow.executeQuery();

         while ( rs.next() ) {
            if ( rs.getInt("eow") > 0 )
               eow = "Y";
         }
      }
      catch ( Exception e ) {
         log.error("exception", e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      rs = null;
      return eow;
   }

   /**
    * Determines whether the customer is charged the under $26K fee
    * @param custid
    * @return String
    */
   private String getUnder26K( String custid )
   {
      String under26k = " ";

      try {
         m_Under26K.setString(1, custid);
         m_Under26K.execute();
         under26k = m_Under26K.getString(2);
      }
      catch ( Exception e ){
         log.error(e);
      }

      return under26k;
   }

   /**
    * Calculates a customer projected sales, it it has been active fewer that 12 months
    * @param activeMo
    * @param sales
    * @return the sales amount
    */
   private double projectedSales( int activeMo, double sales )
   {
      double projSales = sales;

      if ( activeMo < 12 )
         projSales = projSales / activeMo * 12;

      return projSales;
   }

   /**
    * Returns the customer's year to date sales for the sales year
    * @param custid
    * @param salesYear
    * @return the year to date sales
    */
   private double ytdSales( String custid, int salesYear )
   {
      ResultSet rs = null;
      double sales = 0;

      try {
         m_YtdSales.setString(1, custid);
         m_YtdSales.setInt(2, salesYear);

         rs = m_YtdSales.executeQuery();

         while ( rs.next() )
            sales = rs.getDouble("sales");
      }
      
      catch ( Exception e ) {
         log.error("exception", e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }
      
      return sales;
   }
   
   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return XSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, double val)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(val);

      return cell;
   }

   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return XSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, int val)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(val);

      return cell;
   }

   /**
    * Creates a cell of type String
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return XSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, String val)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));

      return cell;
   }

   private void initReport()
   {
      XSSFRow Row = null;

      try {
         if ( m_Spreadsheet ) {
            m_WrkBk = new XSSFWorkbook();
            m_Sheet = m_WrkBk.createSheet();

            Row = m_Sheet.createRow((short)0);

            if ( Row != null ) {
               createCell(Row, 0, "Customer Id");
               createCell(Row, 1, "Name");
               createCell(Row, 2, "Status");
               createCell(Row, 3, "Sales Rep");
               createCell(Row, 4, "Division");
               createCell(Row, 5, "Zip");
               createCell(Row, 6, "Zone");
               createCell(Row, 7, "Active MoC");
               createCell(Row, 8, m_SalesYear + " Sales");
               createCell(Row, 9, "Projected Sales");
               createCell(Row, 10, m_SalesYear + " Acct Sales");
               createCell(Row, 11, "Acct Proj Sales");
               createCell(Row, 12, m_BudgetYear + " Budget");
               createCell(Row, 13, "Top 100");
               createCell(Row, 14, "EOW");
               createCell(Row, 15, "<26K");
               createCell(Row, 16, "Min Order");
               createCell(Row, 17, "BC");
               createCell(Row, 18, "PA");
               createCell(Row, 19, "Total BC/PA");
               createCell(Row, 20, "Rec BC/PA");
               createCell(Row, 21, "Adder Chg");
               createCell(Row, 22, "FRT");
               createCell(Row, 23, "Rec Frt");
            }
         }
         else {
            m_Lines = new StringBuffer();
            m_Lines.append("Customer ID\tName\tStatus\tSales Rep\tDivision\tZip\tZone\tActive Mo\t");
            m_Lines.append(m_SalesYear + " Sales\tProjected Sales\t" + m_SalesYear + " Acct Sales\tAcct Proj Sales\t" + m_BudgetYear + " Budget\t");
            m_Lines.append("Top 100\tEOW\t<26K\tMin Order\tBC\tPA\tTotal BC/PA\tRec BC/PA\tAdder Chg\tFRT\tRec Frt\r\n");
         }
      }

      catch ( Exception e ) {
         log.error(e, e.fillInStackTrace());
      }

      finally {
         Row = null;
      }
   }

   /**
    * Prepares the sql statements
    * 
    * @return boolean - true if statements successfully prepared 
    * @throws Exception
    */
   public boolean prepareStatements() throws Exception
   {
      StringBuffer sql = null;
      boolean prepared = false;
      
      if ( m_OraConn != null ) {      
         sql = new StringBuffer();
         sql.append("declare ");
         sql.append("   custid    varchar2(7) := ?; ");
         sql.append("   year      integer := ?; ");
         sql.append("   acctid    varchar2(7); ");
         sql.append("   setup     date; ");
         sql.append("   months    integer; ");
         sql.append("   setupyr   integer; ");
         sql.append("   sales     number; ");
         sql.append("   projsales number := 0; ");
         sql.append("   totsales  number := 0; ");
         sql.append("   cursor c_stores is ");
         sql.append("      select customer_id ");
         sql.append("      from customer ");
         sql.append("      start with customer_id = acctid ");
         sql.append("      connect by prior customer_id = parent_id; ");
         sql.append("begin ");
         sql.append("   acctid := cust_procs.findtopparent(custid); ");
         sql.append("   for store in c_stores loop ");
         sql.append("      select setup_date into setup from customer where customer_id = store.customer_id; ");
         sql.append("      setupyr := to_number(trim(to_char(setup, 'YYYY'))); ");
         sql.append("      if setupyr < year then ");
         sql.append("         months := 12; ");
         sql.append("      else ");
         sql.append("         months := 12 - to_number(trim(to_char(setup, 'MM'))); ");
         sql.append("      end if; ");
         sql.append("      select sum(nvl(dollars_shipped,0)) into sales from salesyear  ");
         sql.append("      where sales_year = year and ");
         sql.append("            cust_nbr = store.customer_id and ");
         sql.append("            sale_type = 'WAREHOUSE'; ");
         sql.append("      totsales := totsales + nvl(sales,0);  ");
         sql.append("      if months < 12 then ");
         sql.append("         if months = 0 then ");
         sql.append("            sales := 0; ");
         sql.append("         else ");
         sql.append("            sales := sales / months * 12; ");
         sql.append("         end if; ");
         sql.append("      end if; ");
         sql.append("      projsales := round(projsales + nvl(sales,0));  ");
         sql.append("   end loop; ");
         sql.append("   ? := totsales; ");
         sql.append("   ? := projsales; ");
         sql.append("end; ");
         m_AcctSales = (OracleCallableStatement)m_OraConn.prepareCall(sql.toString());
         m_AcctSales.registerOutParameter(3, Types.DOUBLE);
         m_AcctSales.registerOutParameter(4, Types.DOUBLE);
         sql = null;
   
         sql = new StringBuffer();
         sql.append("declare ");
         sql.append("   custid   varchar2(6) := ?; ");
         sql.append("   enddate  date := ?; ");
         sql.append("   mnths    integer; ");
         sql.append("   setup    date; ");
         sql.append("begin ");
         sql.append("   enddate := enddate + 1; ");
         sql.append("   select setup_date into setup from customer where customer_id = custid; ");
         sql.append("   mnths := trunc(months_between(enddate, setup)); ");
         sql.append("   if mnths > 12 then ");
         sql.append("      mnths := 12; ");
         sql.append("   end if; ");
         sql.append("   ? := mnths; ");
         sql.append("end; ");
         m_ActiveMonths = (OracleCallableStatement)m_OraConn.prepareCall(sql.toString());
         m_ActiveMonths.registerOutParameter(3, Types.INTEGER);
         sql = null;
   
         sql = new StringBuffer();
         sql.append("declare ");
         sql.append("   custid    varchar2(7) := ?; ");
         sql.append("   projsales number := ?; ");
         sql.append("   storesales number := ?; ");
         sql.append("   cur_bca   number := -1; ");
         sql.append("   cur_pa    number := -1; ");
         sql.append("   cur_frt   number := -1; ");
         sql.append("   new_bca   number := -1; ");
         sql.append("   new_frt   number := -1; ");
         sql.append("   chg_bca   number := -1; ");
         sql.append("   zipcode   varchar2(10); ");
         sql.append("   frtzone   integer; ");
         sql.append("   totadder  number := -1; ");
         sql.append("begin ");
         sql.append("   cur_bca := cust_procs.adder_pct(custid, 'BASE COST ADDER'); ");
         sql.append("   cur_pa  := cust_procs.adder_pct(custid, 'PERFORMANCE ADDER'); ");
         sql.append("   cur_frt := cust_procs.adder_pct(custid, 'FREIGHT'); ");
         sql.append("   select bc_adder_pct into new_bca from sales_level where min_amt <= projsales and max_amt > projsales; ");
         sql.append("   if new_bca >= 0 then ");
         sql.append("      totadder := cur_bca + cur_pa; ");
         sql.append("   else ");
         sql.append("      totadder := cur_bca; ");
         sql.append("   end if; ");
         sql.append("   chg_bca := new_bca - totadder; ");
         sql.append("   select substr(postal_code, 1, 5) into zipcode from cust_address_view ");
         sql.append("   where customer_id = custid and ");
         sql.append("         addrtype = 'SHIPPING'; ");
         sql.append("   begin ");
         sql.append("      frtzone := transport_procs.get_frt_zone(zipcode); ");
         sql.append("      new_frt := transport_procs.get_frt_rate(zipcode, storesales); ");
         sql.append("   exception ");
         sql.append("      when no_data_found then ");
         sql.append("         null; ");
         sql.append("   end; ");
         sql.append("   ? := cur_bca; ");
         sql.append("   ? := cur_pa; ");
         sql.append("   ? := cur_frt; ");
         sql.append("   ? := chg_bca; ");
         sql.append("   ? := new_frt; ");
         sql.append("   ? := zipcode; ");
         sql.append("   ? := frtzone; ");
         sql.append("   ? := totadder; ");
         sql.append("exception ");
         sql.append("   when others then ");
         sql.append("      null; ");
         sql.append("end; ");
         m_Adders = (OracleCallableStatement)m_OraConn.prepareCall(sql.toString());
         m_Adders.registerOutParameter(4, Types.DOUBLE);
         m_Adders.registerOutParameter(5, Types.DOUBLE);
         m_Adders.registerOutParameter(6, Types.DOUBLE);
         m_Adders.registerOutParameter(7, Types.DOUBLE);
         m_Adders.registerOutParameter(8, Types.DOUBLE);
         m_Adders.registerOutParameter(9, Types.VARCHAR);
         m_Adders.registerOutParameter(10, Types.INTEGER);
         m_Adders.registerOutParameter(11, Types.DOUBLE);
   
         sql = new StringBuffer();
         sql.append("select sum(budget_amount) budget from salesyear ");
         sql.append("where cust_nbr = ? and sales_year = ? and sale_type = 'WAREHOUSE' ");
         m_Budget = m_OraConn.prepareStatement(sql.toString());
         sql = null;
   
         sql = new StringBuffer();
         sql.append("select customer.customer_id, customer.name, customer_status.description status, rep.repname, ");
         sql.append("       rep.division, mkt.class type, min_order, decode(top100.class, 'TOP 100', 'Y', ' ') top100 ");
         sql.append("from customer, customer_status, cust_rep_div_view rep, cust_market_view mkt, cust_market_view top100 ");
         sql.append("where customer.cust_status_id = customer_status.cust_status_id and ");
         sql.append("      customer_status.description <> 'INACTIVE' and ");
         sql.append("      rep.customer_id(+) = customer.customer_id and ");
         sql.append("      rep.rep_type(+) = 'SALES REP' and ");
         sql.append("      mkt.customer_id = customer.customer_id and ");
         sql.append("      mkt.market = 'CUSTOMER TYPE' and ");
         sql.append("      not (mkt.class in ('EMERY','EMPLOYEE','INACTIVE','BACKHAULS','ENAP DS','EMERY DC')) and ");
         sql.append("      top100.customer_id = customer.customer_id and ");
         sql.append("      top100.market = 'TOP 100' ");
         sql.append("order by customer.name ");
         m_Cust = m_OraConn.prepareStatement(sql.toString());
         sql = null;
   
         sql = new StringBuffer();
         sql.append("select count(*) eow ");
         sql.append("from trip_stop_sched, trip_sched ");
         sql.append("where customer_id = ? and");
         sql.append("      trip_stop_sched.ts_id = trip_sched.ts_id and ");
         sql.append("      trip_stop_sched.weeks <> '1,2' ");
         m_Eow = m_OraConn.prepareStatement(sql.toString());
         sql = null;
   
         sql = new StringBuffer();
         sql.append("declare ");
         sql.append("   custid   varchar2(7) := ?; ");
         sql.append("   under26k varchar2(1) := ' '; ");
         sql.append("begin ");
         sql.append("   if service_procs.cust_has_service(custid, service_procs.service_id('UNDER 26K MONTHLY FEE')) then ");
         sql.append("      under26k := 'Y'; ");
         sql.append("   end if; ");
         sql.append("   ? := under26k; ");
         sql.append("end; ");
         m_Under26K = (OracleCallableStatement)m_OraConn.prepareCall(sql.toString());
         m_Under26K.registerOutParameter(2, Types.VARCHAR);
   
         sql = new StringBuffer();
         sql.append("select sum(dollars_shipped) sales from salesyear ");
         sql.append("where cust_nbr = ? and ");
         sql.append("      sales_year = ? and ");
         sql.append("      sale_type = 'WAREHOUSE' ");
         m_YtdSales = m_OraConn.prepareStatement(sql.toString());
         
         prepared = true;
      }
      
      sql = null;
      return prepared;
   }

   /**
    * Sets the parameters for the report. 
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      GregorianCalendar cal = new GregorianCalendar();
      
      m_BudgetYear = Integer.parseInt(params.get(0).value);
      m_SalesYear = Integer.parseInt(params.get(1).value);
      m_Spreadsheet = Boolean.parseBoolean(params.get(2).value);

      m_EndDate = new GregorianCalendar(m_SalesYear, 11, 31);

      if ( m_EndDate.after(cal) )         
         m_EndDate = cal;
      
      cal = null;
   }
}
