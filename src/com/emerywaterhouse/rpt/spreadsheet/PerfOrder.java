/**
 * File: PerfOrder.java
 * Description: perfert order reporting
 *
 * @author Seth Murdock
 *
 * Create Date: 02/14/2007
 * Last Update: $Id: PerfOrder.java,v 1.9 2009/02/18 17:17:50 jfisher Exp $
 * 
 * History
 *    $Log: PerfOrder.java,v $
 *    Revision 1.9  2009/02/18 17:17:50  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.8  2008/10/29 21:45:25  jfisher
 *    Fixed some warnings
 *
 *    Revision 1.7  2008/05/04 05:47:51  smurdock
 *    added selection for Portland and/or Pittston DC's
 *
 *    Revision 1.6  2007/11/16 19:25:15  smurdock
 *    Added code to handle select by customer and select by setup date range
 *
 *    Revision 1.5  2007/02/16 13:57:37  smurdock
 *    added breaks to switch/case in Linotype to stop fallthourgh
 *
 *    Revision 1.4  2007/02/15 17:56:45  smurdock
 *    added sql logic to deal with how item.disp_id's affect silo assignment
 *
 *    Revision 1.3  2007/02/14 16:32:47  jfisher
 *    Fixed formatting issues, unused method warnings, and put in header comments.
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class PerfOrder extends Report 
{
   private  static short BASE_COLS  = 11;
   private int colCnt = BASE_COLS;
   private String m_BegDate;
   private String m_EndDate;
   private Boolean m_by_SetupDate = false;
   private String m_SetupBegDate = "";
   private String m_SetupEndDate = "";
   private String m_Date_Breakdown = "YYYY";
   private String m_VndId = "";
   private String m_CustId = "";
   private String m_Warehouse = "";
   private Boolean m_by_Dept = false;
   private Boolean m_Include_Totals = false;
   private String m_Lines_Units_Dollars = "Lines";
   private PreparedStatement m_PerfOrder;
   private int m_Default_Report = -1;  //-1 means no default report selected
   private String m_Caption = "Perfect Order Report By ";

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private ArrayList<HSSFCellStyle> m_CellStyles = new ArrayList<HSSFCellStyle>();
   
   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;
      
   //
   // Log4j logger
   private Logger m_Log;
       
   
   /**
    * default constructor
    */
   public PerfOrder()
   {      
      super();
      m_Log = Logger.getLogger(RptServer.class);
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
   }
      
   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
      if ( m_CellStyles.size() > 0 ) {
         m_CellStyles.clear();
      }
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyles = null;
            
      super.finalize();
   }
   
   /**
    * generate first part of report header for default reports
    */
   private void LinoTitle()
   {
      switch ( m_Default_Report ) {
         case 0: m_Caption = "Perfect Order Year To Date: "; break;
         case 1: m_Caption = "Perfect Order Year To Date by Dept: "; break;
         case 2: m_Caption = "Perfect Order Year To Date by Day by Dept: "; break;
         case 3: m_Caption = "Perfect Order Month To Date: "; break;
         case 4: m_Caption = "Perfect Order Month To Date by Dept: "; break;
         case 5: m_Caption = "Perfect Order Month To Date by Day by Dept: ";
      }       	   
   }
   
   /**
    * Executes the queries and builds the output file
    * 
    * @return true if the file was built, false if not.
    * @throws FileNotFoundException
    * 
    */  
   private boolean buildOutputFile() throws FileNotFoundException
   {      
      HSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet PerfOrder = null;
      ResultSet ConsolidatedID = null;
      short rowNum = 1;
      boolean result = false;
      float req_cut = 0;
      float allo_cut = 0;
      float ship_cut = 0;
      String date_break = "";
      String dept_num = "00";
      float perfect = 0;
      float partial = 0;
      float silage = 0;
      float p_request = 0;
      float p_allocate = 0;
      float p_ship = 0;
      float req_metric = 0;
      float allo_metric = 0;
      float ship_metric = 0;
      float perfect_metric = 0;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
    	 if (!m_by_Dept)
    		 colCnt = 10;
       
    	 if (!m_Include_Totals)
    		 colCnt = colCnt - 5;
       
         rowNum = createCaptions();         
         if ( m_VndId.length() == 0) {
            m_PerfOrder.setString(1,m_Date_Breakdown);        
            m_PerfOrder.setString(2,m_Date_Breakdown);        
            m_PerfOrder.setString(3,m_Date_Breakdown);
            m_PerfOrder.setString(4,m_BegDate);        
            m_PerfOrder.setString(5,m_EndDate);
            m_PerfOrder.setString(6,m_Date_Breakdown);        
            m_PerfOrder.setString(7,m_Date_Breakdown);        
            m_PerfOrder.setString(8,m_Date_Breakdown);                       
            m_PerfOrder.setString(9,m_BegDate);        
            m_PerfOrder.setString(10,m_EndDate);
            
            if (!m_Lines_Units_Dollars.equals("Lines")){  //we have to account for partials, thus more params in query
               m_PerfOrder.setString(11,m_Date_Breakdown);        
               m_PerfOrder.setString(12,m_Date_Breakdown);                       
               m_PerfOrder.setString(13,m_Date_Breakdown);                       
               m_PerfOrder.setString(14,m_BegDate);        
               m_PerfOrder.setString(15,m_EndDate);
            }
         }
         else {
            m_PerfOrder.setString(1,m_Date_Breakdown);        
            m_PerfOrder.setString(2,m_Date_Breakdown);        
            m_PerfOrder.setString(3,m_Date_Breakdown);
            m_PerfOrder.setString(4,m_VndId);             
            m_PerfOrder.setString(5,m_BegDate);        
            m_PerfOrder.setString(6,m_EndDate);
            m_PerfOrder.setString(7,m_Date_Breakdown);        
            m_PerfOrder.setString(8,m_Date_Breakdown);             
            m_PerfOrder.setString(9,m_Date_Breakdown);                       
            m_PerfOrder.setString(10,m_VndId);             
            m_PerfOrder.setString(11,m_BegDate);        
            m_PerfOrder.setString(12,m_EndDate);

            if (!m_Lines_Units_Dollars.equals("Lines")){  //we have to account for partials, thus more params in query
               m_PerfOrder.setString(13,m_Date_Breakdown);        
               m_PerfOrder.setString(14,m_Date_Breakdown);                       
               m_PerfOrder.setString(15,m_Date_Breakdown);                       
               m_PerfOrder.setString(16,m_VndId);             
               m_PerfOrder.setString(17,m_BegDate);        
               m_PerfOrder.setString(18,m_EndDate);
            }   
         }
        	 
         PerfOrder = m_PerfOrder.executeQuery();

         while ( PerfOrder.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(rowNum);
            date_break = PerfOrder.getString("date_break");
            req_cut = PerfOrder.getFloat("req");
            allo_cut = PerfOrder.getFloat("allo");
            ship_cut = PerfOrder.getFloat("ship");
            perfect = PerfOrder.getFloat("perfect");
            
            if (!m_Lines_Units_Dollars.equals("Lines"))
               partial = PerfOrder.getFloat("partial"); 
            
            silage = req_cut + allo_cut + ship_cut + perfect + partial;
                        
            if (m_by_Dept)
               dept_num = PerfOrder.getString("dept_num");
            
            p_request = silage - req_cut;
            p_allocate = p_request - allo_cut;
            p_ship = p_allocate - ship_cut;
            req_metric = p_request / silage;
            
            if (p_request > 0)
               allo_metric = p_allocate / p_request;
            
            if (p_allocate > 0)
               ship_metric = p_ship / p_allocate;
            
            perfect_metric = (((silage - req_cut)- allo_cut) - ship_cut) / silage;
            
            if (!m_by_Dept){
               row.getCell(0).setCellValue(new HSSFRichTextString(date_break));
               row.getCell(1).setCellValue(req_metric);
               row.getCell(2).setCellValue(allo_metric);
               row.getCell(3).setCellValue(ship_metric);
               row.getCell(4).setCellValue(perfect_metric);
               
               if (m_Include_Totals) {
                  row.getCell(5).setCellValue(req_cut);
                  row.getCell(6).setCellValue(allo_cut);
                  row.getCell(7).setCellValue(ship_cut);
                  row.getCell(8).setCellValue(perfect);
                  row.getCell(9).setCellValue(silage);
               }   
            }
            else{
               row.getCell(0).setCellValue(new HSSFRichTextString(date_break));
               row.getCell(1).setCellValue(new HSSFRichTextString(dept_num));
               row.getCell(2).setCellValue(req_metric);
               row.getCell(3).setCellValue(allo_metric);
               row.getCell(4).setCellValue(ship_metric);
               row.getCell(5).setCellValue(perfect_metric);
               if (m_Include_Totals) {                
                  row.getCell(6).setCellValue(req_cut);
                  row.getCell(7).setCellValue(allo_cut);
                  row.getCell(8).setCellValue(ship_cut);
                  row.getCell(9).setCellValue(perfect);
                  row.getCell(10).setCellValue(silage);
               }	
            }
            rowNum++;            
         }

         m_Wrkbk.write(outFile);
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {         
         row = null;         
         closeRSet(PerfOrder);
         closeRSet(ConsolidatedID);                 
         try {
            outFile.close();
         }

         catch( Exception e ) {
            m_Log.error("exception:", e);
         }

         outFile = null;
      }
      
      return result;
   }
   
   /**
    * Builds the sql based on the type of filter requested by the user.
    * A union of two queries: the first gets the lines that were not perfect
    * (i.e they show up in item_cut_reason) and the second gets the perfect lines.
    * @return A complete sql statement.
    */
   private String buildSql()
   {
	   StringBuffer sql = new StringBuffer(256);
	  
      sql.append ("select date_break,  ");//bb
      if (!m_by_Dept){      
         sql.append ("max(decode(silo_id,0, silocut,0)) req, "); //nn
         sql.append ("max(decode(silo_id,1, silocut,0)) allo, "); //nn
         sql.append ("max(decode(silo_id,2, silocut,0)) ship, "); //nn
         sql.append ("max(decode(silo_id,100, silocut,0)) perfect, "); //nn
         sql.append ("sum(cutto) totie "); //nn
      }
      else {      
         sql.append (" dept_num, ");//dd
         sql.append ("max(decode(silo_id,0, cutmo,0)) req, "); //nn
         sql.append ("max(decode(silo_id,1, cutmo,0)) allo, ");//dd
         sql.append ("max(decode(silo_id,2, cutmo,0)) ship, ");//dd
         sql.append ("max(decode(silo_id,100, cutmo,0)) perfect ");//dd
         sql.append ("from ");//dd
         sql.append ("  (select distinct(date_break), dept_num, silo_id, ");//dd
         sql.append ("   sum (cutto) over (partition by date_break, silo_id, dept_num) cutmo ");//dd
      }
      sql.append    ("   from ");//bb
      sql.append    ("    (select distinct(to_char(id.invoice_date,?)) Date_Break, ");//bb //Param 1 date_break
      sql.append    ("     icr.cutcode, silo_id, icr.disp_id,  ");//bb
      
      if (!m_by_Dept){      
         sql.append ("     count(distinct(icr.inv_dtl_id)) over(partition by cs.silo_id, to_char(id.invoice_date,?)) silocut, ");  //pram 2nn date_break//nn
         sql.append ("     count(distinct(icr.inv_dtl_id)) over(partition by icr.disp_id, icr.cutcode, to_char(id.invoice_date,?)) cutto  ");     //param 3nn date_break     //nn
      }
      else{
         sql.append ("     dept_num, ");//dd
         sql.append ("     count(distinct(icr.inv_dtl_id)) over(partition by cs.silo_id, ed.dept_num, to_char(id.invoice_date,?)) silocut, ");//parm 2dd date_break//dd
         sql.append ("     count(distinct(icr.inv_dtl_id)) over(partition by icr.disp_id, icr.cutcode, ed.dept_num, to_char(id.invoice_date, ?)) cutto     ");//parm 3dd date break//dd
      }
      sql.append    ("     from item_cut_reason icr, cutcode_silo cs, inv_dtl id ");//bb
      
      //Department and/or setup date needs table item to filter
      if ((m_by_Dept) && !(m_by_SetupDate))      
         sql.append ("     ,item, emery_dept ed  ");//dd
      else if ((m_by_Dept) && (m_by_SetupDate))
          sql.append ("     ,item,emery_dept ed  ");//dd??????
      else if (!(m_by_Dept) && (m_by_SetupDate))
          sql.append ("     ,item  ");//dd??????
   	  
      sql.append    ("     where icr.cutcode is not null ");//bb
      sql.append    ("     and icr.cutcode = cs.cutcode ");//bb
      sql.append    ("     and icr.disp_id = cs.disp_id ");//bb
      sql.append    ("     and icr.inv_dtl_id = id.inv_dtl_id ");//bb
      if (m_Warehouse != null && m_Warehouse.length() > 0){
         sql.append    ("     and id.warehouse = '");
         sql.append(m_Warehouse);
         sql.append("' ");
      }      
      
      if ( m_VndId != null && m_VndId.length() > 0 )
          sql.append("     and id.vendor_nbr =  ? ");                   //param 4bb if byvendor          
      if ( m_CustId != null && m_CustId.length() > 0 ){
          sql.append("     and id.cust_nbr =  ");
          sql.append(m_CustId);
          sql.append(" ");
      }

      sql.append    ("     and id.invoice_date >= to_date(?,'mm/dd/yyyy')  ");             //param 4bb start_date 5bb if byvendor");//bb
      sql.append    ("     and id.invoice_date <= to_date(?,'mm/dd/yyyy')  ");             //param 5bb end_date 6bb if byvendor");//bb
      
      if ((m_by_Dept) && !(m_by_SetupDate)){      
         sql.append ("     and id.item_nbr = item.item_id ");//dd
         sql.append ("     and item.dept_id= ed.dept_id ");//dd
      }
      else if ((m_by_Dept) && (m_by_SetupDate)){
          sql.append ("     and id.item_nbr = item.item_id ");//dd
          sql.append ("     and item.dept_id= ed.dept_id ");//dd
          sql.append ("     and item.setup_date >= to_date('");//dd??
          sql.append(m_SetupBegDate);
          sql.append ("','mm/dd/yyyy') ");//dd??
          
          sql.append ("     and item.setup_date <=  to_date('");//dd??
          sql.append(m_SetupEndDate);
          sql.append ("','mm/dd/yyyy') ");//dd??
          sql.append("  ");
        
    	  
      }
      else if (!(m_by_Dept) && (m_by_SetupDate)){
          sql.append ("     and id.item_nbr = item.item_id ");//dd
          sql.append ("     and item.setup_date >=  to_date('");//dd??
          sql.append(m_SetupBegDate);
          sql.append ("','mm/dd/yyyy') ");//dd??
          sql.append ("     and item.setup_date <=  to_date('");//dd??
          sql.append(m_SetupEndDate);
          sql.append ("','mm/dd/yyyy') ");//dd??
          sql.append("  ");
    	  
      }
      
      sql.append    ("     union ");//bb
      sql.append    ("     select  ");//bb //this query gets the results that were perfect, e.g not in
      sql.append    ("     distinct(to_char(id.invoice_date,?)) date_break,  'PERFECT','100', 1, ");//parm 6 bb date_break 7bb if by vendor //bb
      if (!m_by_Dept){      
         sql.append ("     count(distinct(inv_dtl_id)) over(partition by to_char(id.invoice_date,?)) silocut, ");                 //param 7nn date_break 8nn if byvendor//nn
         sql.append ("     count(distinct(inv_dtl_id)) over(partition by to_char(id.invoice_date,?)) cutto  ");                   //param 8nn date_break 9nn if byvendor//nn
      }
      else {
         sql.append ("     dept_num, ");//dd
         sql.append ("     count(distinct(inv_dtl_id)) over(partition by ed.dept_num, to_char(id.invoice_date,?)) silocut,  ");//param 7 dd date_break 8dd if byvendor//dd
         sql.append ("     count(distinct(inv_dtl_id)) over(partition by ed.dept_num, to_char(id.invoice_date,?)) cutto ");//param 8 dd date_break 9dd if byvendor//dd
      }
      sql.append    ("     from inv_dtl id ");//bb
      
      if ((m_by_Dept) && !(m_by_SetupDate)){
          sql.append ("    , item, emery_dept ed   ");//    	  
      }
      else if ((m_by_Dept) && (m_by_SetupDate)){
          sql.append ("    , item, emery_dept ed   ");//   	  
      }
      else if (!(m_by_Dept) && (m_by_SetupDate)){
          sql.append ("    , item   ");//   	  
      }
    // if (m_by_Dept)      
    //     sql.append ("    , item, emery_dept ed   ");//
      sql.append    ("     where inv_dtl_id >= 50535808  ");//bb
      sql.append    ("     and id.tran_type = 'SALE'  ");//bb
      sql.append    ("     and id.sale_type = 'WAREHOUSE'   ");//bb
      
      if ( m_VndId != null && m_VndId.length() > 0 )
         sql.append ("     and id.vendor_nbr =  ? ");                   //param 10 bb if byvendor

      if ( m_CustId != null && m_CustId.length() > 0 ){
          sql.append("     and id.cust_nbr =  ");
          sql.append(m_CustId);
          sql.append(" ");
      }
      
      if (m_Warehouse != null && m_Warehouse.length() > 0){
          sql.append    ("     and id.warehouse = '");
          sql.append(m_Warehouse);
          sql.append("' ");          
       }      
      
      sql.append    ("     and id.invoice_date >= to_date(?,'mm/dd/yyyy') "); //param 9 bb start date 11bb if byvendor//bb
      sql.append    ("     and id.invoice_date <= to_date(?,'mm/dd/yyyy')  ");//param 10 bb end_date 12 bb if byvendor//bb
      
      if ((m_by_Dept) && !(m_by_SetupDate)){
          sql.append ("     and id.item_nbr = item.item_id ");//dd
          sql.append ("     and item.dept_id = ed.dept_id ");//dd
       }
      else if ((m_by_Dept) && (m_by_SetupDate)){
          sql.append ("     and id.item_nbr = item.item_id ");//dd
          sql.append ("     and item.dept_id = ed.dept_id ");//dd
          sql.append ("     and item.setup_date >= to_date('");//dd
          sql.append(m_SetupBegDate);
          sql.append ("','mm/dd/yyyy') ");//dd??


          sql.append ("     and item.setup_date <= to_date('");//dd
          sql.append(m_SetupEndDate);
          sql.append ("','mm/dd/yyyy') ");//dd??

          sql.append("  ");
      }
      else if (!(m_by_Dept) && (m_by_SetupDate)){
          sql.append ("     and id.item_nbr = item.item_id ");//dd
          sql.append ("     and item.setup_date >=  to_date('");//dd
          sql.append(m_SetupBegDate);
          sql.append ("','mm/dd/yyyy') ");//dd??
          sql.append ("     and item.setup_date <=   to_date('");//dd
          sql.append(m_SetupEndDate);
          sql.append ("','mm/dd/yyyy') ");//dd??
          sql.append("  ");
    	  
      }
      //if (m_by_Dept) {      
        // sql.append ("     and id.item_nbr = item.item_id ");//dd
        // sql.append ("     and item.dept_id = ed.dept_id ");//dd
      //}
      sql.append    ("     and inv_dtl_id not in  ");//bb
      sql.append    ("       (select inv_dtl_id from item_cut_reason))");//bb
      
      if (m_by_Dept){
         sql.append(")");      
         sql.append ("group by date_break, dept_num ");//dd
      }
      else 	//
         sql.append ("group by date_break ");//nn
	  
         
      return sql.toString();
   }

   /**
    * Reports by units and dollars include partial shipments on a line.
    * In the line based reports a line is either perfect or not, but to
    * get the proper unit or dollar totals we need to include partial line shipments
    */
   private String buildSql_with_partials()
   {   
	  StringBuffer sql = new StringBuffer(256);
	  
	  sql.append ("select date_break,  ");//bb
	  
     if (!m_by_Dept){      
		 sql.append ("max(decode(silo_id,0, total_silo,0)) req, "); //nn
		 sql.append ("max(decode(silo_id,1, total_silo,0)) allo, "); //nn
	     sql.append ("max(decode(silo_id,2, total_silo,0)) ship, "); //nn
	     sql.append ("max(decode(silo_id,100, total_silo,0)) perfect, "); //nn
	     sql.append ("max(decode(silo_id,101, total_silo,0)) partial, "); //nn
	     sql.append ("sum(total_outcome) totie "); //nn
	  }
	  else {      
	     sql.append (" dept_num, ");//dd
		  sql.append ("max(decode(silo_id,0, cutmo,0)) req, "); //nn
	     sql.append ("max(decode(silo_id,1, cutmo,0)) allo, ");//dd
	     sql.append ("max(decode(silo_id,2, cutmo,0)) ship, ");//dd
	     sql.append ("max(decode(silo_id,100, cutmo,0)) perfect, ");//dd
	     sql.append ("max(decode(silo_id,101, cutmo,0)) partial "); //nn
	     sql.append (" from ");//dd
	     sql.append ("   (select distinct(date_break), dept_num, silo_id, disp_id, ");//dd
	     sql.append ("   sum (total_outcome) over (partition by date_break, silo_id, dept_num) cutmo ");//dd
	  }
	  
     sql.append    ("   from ");//bb
	  sql.append    ("     (select distinct(to_char(id.invoice_date,?)) Date_Break, ");//bb //Param 1 date_break
	  sql.append    ("     icr.cutcode, silo_id, icr.disp_id,  ");//bb
	      
	  if (!m_by_Dept){
	     if (m_Lines_Units_Dollars.equals("Units")){ 
	        sql.append ("  sum(sicut) over(partition by cs.silo_id, to_char(id.invoice_date,?)) total_silo, ");  //pram 2nn date_break//nn
	        sql.append ("  sum(sicut) over(partition by icr.disp_id, icr.cutcode, to_char(id.invoice_date,?)) total_outcome  ");     //param 3nn date_break     //nn
	     }          
	     else{  //report by dollars
		    sql.append ("  sum((qty_ordered - qty_shipped) * unit_sell) over(partition by cs.silo_id, to_char(id.invoice_date,?)) total_silo, ");  //pram 2nn date_break//nn
		    sql.append ("  sum((qty_ordered - qty_shipped) * unit_sell) over(partition by icr.disp_id, icr.cutcode, to_char(id.invoice_date,?)) total_outcome  ");     //param 3nn date_break     //nn	    		 
	     }
	  }	 
	  else{ //by department
	     sql.append    ("  dept_num, ");//dd
	     if (m_Lines_Units_Dollars.equals("Units")){ 
   	        sql.append ("  sum(sicut) over(partition by cs.silo_id, ed.dept_num, to_char(id.invoice_date,?)) total_silo, ");//parm 2dd date_break//dd
	        sql.append ("  sum(sicut) over(partition by icr.disp_id, icr.cutcode, ed.dept_num, to_char(id.invoice_date, ?)) total_outcome     ");//parm 2dd date break//dd
	     }
	     else { //report by dollars
	   	    sql.append ("   sum((qty_ordered - qty_shipped) * unit_sell) over(partition by cs.silo_id, ed.dept_num, to_char(id.invoice_date,?)) total_silo, ");//parm 2dd date_break//dd
	        sql.append ("   sum((qty_ordered - qty_shipped) * unit_sell) over(partition by icr.disp_id, icr.cutcode, ed.dept_num, to_char(id.invoice_date, ?)) total_outcome     ");//parm 2dd date break//dd
	     }	 
	  }
	  sql.append       ("   from ");           //item_cut_reason icr, "
	  sql.append       ("      (select inv_dtl_id, item_id, cutcode, disp_id, "); //this line and 3 following deal with duplicate icr entries caused by shipment merges
	  sql.append       ("      max(decode(inv_dtl_id,inv_dtl_id,sicut,null)) sicut ");
	  sql.append       ("         from item_cut_reason ");
	  sql.append       ("   group by inv_dtl_id, item_id, cutcode, disp_id) icr, "); 
	  sql.append       ("   cutcode_silo cs, inv_dtl id ");//bb
	  
      if ((m_by_Dept) && !(m_by_SetupDate)){
          sql.append ("    , item, emery_dept ed   ");//    	  
      }
      else if ((m_by_Dept) && (m_by_SetupDate)){
          sql.append ("    , item, emery_dept ed   ");//   	  
      }
      else if (!(m_by_Dept) && (m_by_SetupDate)){
          sql.append ("    , item   ");//   	  
      }

	  //if (m_by_Dept)      
	     //sql.append    ("   ,item, emery_dept ed  ");//dd
	  
     sql.append       ("   where icr.cutcode is not null ");//bb
	  sql.append       ("   and icr.cutcode = cs.cutcode ");//bb
	  sql.append       ("   and icr.disp_id = cs.disp_id ");//bb
	  sql.append       ("   and icr.inv_dtl_id = id.inv_dtl_id ");//bb
	  
     if ( m_VndId != null && m_VndId.length() > 0 )
	     sql.append(" and id.vendor_nbr =  ? ");                   //param 4bb if byvendor
     if ( m_CustId != null && m_CustId.length() > 0 ){
         sql.append("     and id.cust_nbr =  ");
         sql.append(m_CustId);
         sql.append(" ");
     }
     if (m_Warehouse != null && m_Warehouse.length() > 0){
         sql.append    ("     and id.warehouse = '");
         sql.append(m_Warehouse);
         sql.append("' ");          
      }      
     
      sql.append       ("   and invoice_date >= to_date(?,'mm/dd/yyyy')  ");             //param 4bb start_date 5bb if byvendor");//bb
	  sql.append       ("   and invoice_date <= to_date(?,'mm/dd/yyyy')  ");             //param 5bb end_date 6bb if byvendor");//bb 
	  
	     if ((m_by_Dept) && !(m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.dept_id = ed.dept_id ");//dd
	       }
	      else if ((m_by_Dept) && (m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.dept_id = ed.dept_id ");//dd
	          sql.append ("     and item.setup_date >= to_date('");//dd
	          sql.append(m_SetupBegDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??


	          sql.append ("     and item.setup_date <= to_date('");//dd
	          sql.append(m_SetupEndDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??

	          sql.append("  ");
	      }
	      else if (!(m_by_Dept) && (m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.setup_date >=  to_date('");//dd
	          sql.append(m_SetupBegDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??
	          sql.append ("     and item.setup_date <=   to_date('");//dd
	          sql.append(m_SetupEndDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??
	          sql.append("  ");
	    	  
	      }
     //if (m_by_Dept){      
	     //sql.append    ("   and id.item_nbr = item.item_id ");//dd
	     //sql.append    ("   and item.dept_id= ed.dept_id ");//dd
	  //}
	  sql.append       ("union ");//bb
	  sql.append       ("   select  ");//bb
	  sql.append       ("   distinct(to_char(id.invoice_date,?)) date_break,  'PERFECT','100', 1, ");//parm 6 bb date_break 7bb if byvendor //bb
	  if (!m_by_Dept){
		 if (m_Lines_Units_Dollars.equals("Units")){ 
	        sql.append ("   sum(qty_shipped) over(partition by to_char(id.invoice_date,?)) total_silo, ");                 //param 7nn date_break 8nn if byvendor//nn
	        sql.append ("   sum(qty_shipped) over(partition by to_char(id.invoice_date,?)) total_outcome  ");                   //param 8nn date_break 9nn if byvendor//nn
	     }
		 else {  //report by dollars
		    sql.append ("   sum(qty_shipped * unit_sell) over(partition by to_char(id.invoice_date,?)) total_silo, ");                 //param 7nn date_break 8nn if by vendor//nn
		    sql.append ("   sum(qty_shipped * unit_sell) over(partition by to_char(id.invoice_date,?)) total_outcome  ");                   //param 8nn date_break 9nn if by vendor//nn		    	 
		 }
	  }   
	  else { //by department
	     sql.append    ("   dept_num, ");//dd
		 if (m_Lines_Units_Dollars.equals("Units")){ 
	        sql.append ("   sum(qty_shipped) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_silo,  ");//param 7 dd date_break 8dd if byvendor//dd
	        sql.append ("   sum(qty_shipped) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_outcome ");//param 8 dd date_break 9dd if byvendor//dd
		 }
		 else {
		    sql.append ("   sum(qty_shipped * unit_sell) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_silo,  ");//param 7 dd date_break 8dd if byvendor//dd
		    sql.append ("   sum(qty_shipped * unit_sell) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_outcome ");//param 8 dd date_break 9dd if byvendor//dd
		 }
     }
	  
     sql.append       ("   from inv_dtl id ");//bb
	  
     if ((m_by_Dept) && !(m_by_SetupDate)){
         sql.append ("    , item, emery_dept ed   ");//    	  
     }
     else if ((m_by_Dept) && (m_by_SetupDate)){
         sql.append ("    , item, emery_dept ed   ");//   	  
     }
     else if (!(m_by_Dept) && (m_by_SetupDate)){
         sql.append ("    , item   ");//   	  
     }

//     if (m_by_Dept)      
	    // sql.append    ("   , item, emery_dept ed  ");//
	  
      sql.append       ("      where inv_dtl_id >= 50535808  ");//bb
	  sql.append       ("      and id.tran_type = 'SALE'  ");//bb
	  sql.append       ("      and id.sale_type = 'WAREHOUSE'   ");//bb
	 if (m_Warehouse != null && m_Warehouse.length() > 0){
	    sql.append    ("     and id.warehouse = '");
	    sql.append(m_Warehouse);
	    sql.append("' ");          
	 }      
	  
     if ( m_VndId != null && m_VndId.length() > 0 )
	     sql.append    ("      and id.vendor_nbr =  ? ");                   //param 10bb if by vendor
     if ( m_CustId != null && m_CustId.length() > 0 ){
         sql.append("     and id.cust_nbr =  ");
         sql.append(m_CustId);
         sql.append(" ");
     }
      sql.append       ("      and invoice_date >= to_date(?,'mm/dd/yyyy') "); //param 9 bb start date 11 bb if byvendor//bb
	  sql.append       ("      and invoice_date <= to_date(?,'mm/dd/yyyy')  ");//param 10 bb end_date 12 bb if byvendor //bb
	  
	     if ((m_by_Dept) && !(m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.dept_id = ed.dept_id ");//dd
	       }
	      else if ((m_by_Dept) && (m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.dept_id = ed.dept_id ");//dd
	          sql.append ("     and item.setup_date >= to_date('");//dd
	          sql.append(m_SetupBegDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??


	          sql.append ("     and item.setup_date <= to_date('");//dd
	          sql.append(m_SetupEndDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??

	          sql.append("  ");
	      }
	      else if (!(m_by_Dept) && (m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.setup_date >=  to_date('");//dd
	          sql.append(m_SetupBegDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??
	          sql.append ("     and item.setup_date <=   to_date('");//dd
	          sql.append(m_SetupEndDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??
	          sql.append("  ");
	    	  
	      }
   // if (m_by_Dept) {      
	     //sql.append    ("      and id.item_nbr = item.item_id ");//dd
	     //sql.append    ("      and item.dept_id= ed.dept_id ");//dd
	  //}
	  
     sql.append       ("      and inv_dtl_id not in  ");//bb
	  sql.append       ("      (select inv_dtl_id from item_cut_reason)");//bb
	  sql.append       ("   union ");//bb
	  sql.append       ("   select  ");//bb
	  sql.append       ("   distinct(to_char(id.invoice_date,?)) date_break,  'PARTIAL','101', 1,  ");//parm 11 bb date_break 13 bb if byvendor//bb
	  if (!m_by_Dept){
	     if (m_Lines_Units_Dollars.equals("Units")){ 
	        sql.append ("   sum(qty_shipped) over(partition by to_char(id.invoice_date,?)) total_silo, ");                 //param 12nn date_break 14nn if byvendor//nn
	        sql.append ("   sum(qty_shipped) over(partition by to_char(id.invoice_date,?)) total_outcome  ");                   //param 13nn date_break 15nn if byvendor//nn
	     }
	     else{ //report by dollars
		    sql.append ("   sum(qty_shipped * unit_sell) over(partition by to_char(id.invoice_date,?)) total_silo, ");                 //param 12nn date_break 14nn if byvendor//nn
		    sql.append ("   sum(qty_shipped * unit_sell) over(partition by to_char(id.invoice_date,?)) total_outcome  ");                   //param 13nn date_break 15nn if byvendor//nn
	     }
	  }
	  else {
	     sql.append ("     dept_num, ");//dd
		 
        if (m_Lines_Units_Dollars.equals("Units")){ 	         
	        sql.append ("      sum(qty_shipped) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_silo,  ");//param 12 dd date_break 14dd if byvendor//dd
	        sql.append ("      sum(qty_shipped) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_outcome ");//param 13 dd date_break 15dd if byvendor//dd
		 }
		 else{
		    sql.append ("      sum(qty_shipped * unit_sell) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_silo,  ");//param 12 dd date_break 14dd if byvendor//dd
		    sql.append ("      sum(qty_shipped * unit_sell) over(partition by ed.dept_num, to_char(id.invoice_date,?)) total_outcome ");//param 13 dd date_break 15dd if byvendor//dd				 
		 }
      }
	  sql.append       ("      from inv_dtl id ");//bb
      if ((m_by_Dept) && !(m_by_SetupDate)){
          sql.append ("    , item, emery_dept ed   ");//    	  
      }
      else if ((m_by_Dept) && (m_by_SetupDate)){
          sql.append ("    , item, emery_dept ed   ");//   	  
      }
      else if (!(m_by_Dept) && (m_by_SetupDate)){
          sql.append ("    , item   ");//   	  
      }
	  
    // if (m_by_Dept)      
	     //sql.append    ("    , item, emery_dept ed   ");//
	  sql.append       ("      where inv_dtl_id >= 50535808  ");//bb
	  sql.append       ("      and id.tran_type = 'SALE'  ");//bb
	  sql.append       ("      and id.sale_type = 'WAREHOUSE'   ");//bb
	  if (m_Warehouse != null && m_Warehouse.length() > 0){
	     sql.append    ("     and id.warehouse = '");
	     sql.append(m_Warehouse);
	     sql.append("' ");          
	  }      
	  
     if ( m_VndId != null && m_VndId.length() > 0 )
	     sql.append    ("      and id.vendor_nbr =  ? ");                   //param 16bb if by vendor
     if ( m_CustId != null && m_CustId.length() > 0 ){
         sql.append("     and id.cust_nbr =  ");
         sql.append(m_CustId);
         sql.append(" ");
     }

      sql.append       ("      and invoice_date >= to_date(?,'mm/dd/yyyy') "); //param 14 bb start date 17 bb if byvendor//bb
	  sql.append       ("      and invoice_date <= to_date(?,'mm/dd/yyyy')  ");//param 15 bb end_date 18 bb if byvendor//bb
	  
	     if ((m_by_Dept) && !(m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.dept_id = ed.dept_id ");//dd
	       }
	      else if ((m_by_Dept) && (m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.dept_id = ed.dept_id ");//dd
	          sql.append ("     and item.setup_date >= to_date('");//dd
	          sql.append(m_SetupBegDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??


	          sql.append ("     and item.setup_date <= to_date('");//dd
	          sql.append(m_SetupEndDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??

	          sql.append("  ");
	      }
	      else if (!(m_by_Dept) && (m_by_SetupDate)){
	          sql.append ("     and id.item_nbr = item.item_id ");//dd
	          sql.append ("     and item.setup_date >=  to_date('");//dd
	          sql.append(m_SetupBegDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??
	          sql.append ("     and item.setup_date <=   to_date('");//dd
	          sql.append(m_SetupEndDate);
	          sql.append ("','mm/dd/yyyy') ");//dd??
	          sql.append("  ");
	    	  
	      }
     //if (m_by_Dept) {      
	    // sql.append    ("      and id.item_nbr = item.item_id ");//dd
	     //sql.append    ("      and item.dept_id= ed.dept_id ");//dd
	  //}
	  
     sql.append       ("      and inv_dtl_id  in  ");//bb
	  sql.append       ("      (select inv_dtl_id from item_cut_reason))");//bb
	  
     if (m_by_Dept){
	     sql.append(")");      
	     sql.append    ("   group by date_break, dept_num ");//dd
	  }
	  else 	
	     sql.append    ("   group by date_break ");//nn
        
      return sql.toString();
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_PerfOrder);
   }
   
   /**
    * Creates the report title and the captions.
    */
   private short createCaptions() {
      HSSFFont fontTitle;
      HSSFCellStyle styleTitle;   // Bold, centered      
      HSSFCellStyle styleTitleLeft;   // Bold, Left Justified      
      HSSFCellStyle styleTitleRight;   // Bold, Right Justified      
      HSSFRow row = null;
      HSSFCell cell = null;
      short rowNum = 0;
      StringBuffer caption = new StringBuffer(256);
      
      if (m_Default_Report >= 0){  //report header based on preset report selection
         LinoTitle();
         caption.append(m_Caption);
      }
      else {                       // no preset report so build header based on selection criteria
         caption.append(m_Caption);
         caption.append(m_Lines_Units_Dollars);
         caption.append(": ");
      }
      
      if ( m_Sheet == null )
         return 0;
      
      fontTitle = m_Wrkbk.createFont();
      fontTitle.setFontHeightInPoints((short)10);
      fontTitle.setFontName("Arial");
      fontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
      
      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(fontTitle);
      styleTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      
      styleTitleLeft = m_Wrkbk.createCellStyle();
      styleTitleLeft.setFont(fontTitle);
      styleTitleLeft.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      
      styleTitleRight = m_Wrkbk.createCellStyle();
      styleTitleRight.setFont(fontTitle);
      styleTitleRight.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      
      //
      // set the report title
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(styleTitleLeft);
       
      caption.append(m_BegDate);
      caption.append(" - ");
      caption.append(m_EndDate);
      
      if (m_Warehouse != null && m_Warehouse.length() > 0){
          caption.append(" ");
          caption.append(m_Warehouse);
          caption.append(" ");
       }      
      
      if (m_VndId != null && m_VndId.length() > 0){
         caption.append(" Vendor: ");
         caption.append(m_VndId);
      }
      if ( m_CustId != null && m_CustId.length() > 0 ){
          caption.append(" Customer: ");
          caption.append(m_CustId);
       }
      if (m_by_SetupDate) {
          caption.append(" Setup Date: ");
          caption.append(m_SetupBegDate);
          caption.append(" - ");
          caption.append(m_SetupEndDate);           
      }

      cell.setCellValue(new HSSFRichTextString(caption.toString()));
      
      rowNum ++;
      rowNum ++;
      row = m_Sheet.createRow(rowNum);

      try {
         if ( row != null ) {
            for ( int i = 0; i < colCnt; i++ ) {
               cell = row.createCell(i);
               if (i < 2)
                  cell.setCellStyle(styleTitleLeft);
               else
                  cell.setCellStyle(styleTitleRight);
           	   
            }
            if (!m_by_Dept){
               row.getCell(0).setCellValue(new HSSFRichTextString("Date"));
               m_Sheet.setColumnWidth(0, 3000);
               row.getCell(1).setCellValue(new HSSFRichTextString("Request"));    
               row.getCell(2).setCellValue(new HSSFRichTextString("Plan"));
               row.getCell(3).setCellValue(new HSSFRichTextString("Ship"));
               row.getCell(4).setCellValue(new HSSFRichTextString("Perfect"));
               
               if (m_Include_Totals){
                  row.getCell(5).setCellValue(new HSSFRichTextString("Req Cut"));
                  row.getCell(6).setCellValue(new HSSFRichTextString("Plan Cut"));
                  row.getCell(7).setCellValue(new HSSFRichTextString("Ship Cut"));
                  row.getCell(8).setCellValue(new HSSFRichTextString("Perf Line"));
                  row.getCell(9).setCellValue(new HSSFRichTextString("Total"));
                  
                  if (m_Lines_Units_Dollars.equals("Dollars")){
                     m_Sheet.setColumnWidth(5, 3500);
                     m_Sheet.setColumnWidth(6, 3500);
                     m_Sheet.setColumnWidth(7, 3500);
                     m_Sheet.setColumnWidth(8, 3500);
                     m_Sheet.setColumnWidth(9, 3500);
                  }
               }   
            }
            else {
               row.getCell(0).setCellValue(new HSSFRichTextString("Date"));
               m_Sheet.setColumnWidth(0, 3000);
               row.getCell(1).setCellValue(new HSSFRichTextString("Dept"));                   
               row.getCell(2).setCellValue(new HSSFRichTextString("Request"));    
               row.getCell(3).setCellValue(new HSSFRichTextString("Plan"));
               row.getCell(4).setCellValue(new HSSFRichTextString("Ship"));
               row.getCell(5).setCellValue(new HSSFRichTextString("Perfect"));
    
               if (m_Include_Totals){                
                  row.getCell(6).setCellValue(new HSSFRichTextString("Req Cut"));
                  row.getCell(7).setCellValue(new HSSFRichTextString("Plan Cut"));
                  row.getCell(8).setCellValue(new HSSFRichTextString("Ship Cut"));
                  row.getCell(9).setCellValue(new HSSFRichTextString("Perf Line"));
                  row.getCell(10).setCellValue(new HSSFRichTextString("Total"));
               
                  if (m_Lines_Units_Dollars.equals("Dollars")){
                     m_Sheet.setColumnWidth(6, 3500);
                     m_Sheet.setColumnWidth(7, 3500);
                     m_Sheet.setColumnWidth(8, 3500);
                     m_Sheet.setColumnWidth(9, 3500);
                     m_Sheet.setColumnWidth(10, 3500);
                  }                 
               }            	
            }
         }
      }
      
      finally {
         row = null;
         cell = null;
         fontTitle = null;
         styleTitle = null;
         caption = null;
      }
            
      return ++rowNum;
   }
   
   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    * 
    * @return The fromatted row of the spreadsheet.
    */
   private HSSFRow createRow(short rowNum) 
   {   
      HSSFRow row = null;
      HSSFCell cell = null;
      
      if ( m_Sheet == null )
         return row;

      row = m_Sheet.createRow(rowNum);

      //
      // set the type and style of the cell.
      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {            
            cell = row.createCell(i);
            cell.setCellStyle(m_CellStyles.get(i));
         }
      }

      return row;
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         setupWorkbook();
         prepareStatements();      
         created = buildOutputFile();            
      }
      
      catch ( Exception ex ) {
         m_Log.fatal("exception:", ex);
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
    *     
    */
   private void prepareStatements() throws Exception 
   {            
      if ( m_OraConn != null ) {		  
         // our main big honking query gets put together in buildsql	  
         if (m_Lines_Units_Dollars.equals("Lines"))
            m_PerfOrder = m_OraConn.prepareStatement(buildSql());
         else
 		      m_PerfOrder = m_OraConn.prepareStatement(buildSql_with_partials());       	  
	  }
   }
      
   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    *
    */
   public void setParams(ArrayList<Param> params) {
	   
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;
      
       
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                            
         if ( param.name.equals("breakdown") ) 
            if (param.value.equals("Month"))
               m_Date_Breakdown = "YYYY/MM";    
        	else if (param.value.equals("Day"))
        	   m_Date_Breakdown = "YYYY/MM/DD";  //otherwise, default to "Year" "YYYY' 

         if ( param.name.equals("dc") ) 
            if (param.value.equals("01"))
               m_Warehouse = "PORTLAND";    
         	else if (param.value.equals("04"))
         	   m_Warehouse = "PITTSTON";  //otherwise, default to "" 
           
         
         if ( param.name.equals("Unit") )
          	m_Lines_Units_Dollars = param.value;
         
         if ( param.name.equals("rptIndex") )
          	m_Default_Report = Integer.parseInt(param.value);
         
         if ( param.name.equals("vendor") && param.value.trim().length() > 0 )
             m_VndId = (param.value);

         if ( param.name.equals("customer") && param.value.trim().length() > 0 )
             m_CustId = (param.value);
                  
         if ( param.name.equals("dept") )
        	m_by_Dept = Boolean.parseBoolean(param.value);
         
         if ( param.name.equals("totals") )
         	m_Include_Totals = Boolean.parseBoolean(param.value);         

         if ( param.name.equals("begdate") ) 
             m_BegDate = param.value;
           
          if ( param.name.equals("enddate") )
             m_EndDate = param.value;
          
          if ( param.name.equals("setupbegdate") ){       	  
              m_SetupBegDate = param.value;
        	  m_by_SetupDate = true;       	  
          }
            
           if ( param.name.equals("setupenddate") ){
              m_SetupEndDate = param.value;
        	  m_by_SetupDate = true;       	                
           }
      }
      
      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("pr.xls");
      m_FileNames.add(fname.toString());
   }
   
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      HSSFCellStyle styleText;      // Text right justified
      HSSFCellStyle styleInt;       // Style with 0 decimals
      HSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00) 
      HSSFCellStyle stylePct;       // Style with 0 decimals + %
      styleText = m_Wrkbk.createCellStyle();
      //styleText.setFont(m_FontData);
      styleText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      styleMoney.setDataFormat((short)8);
      
      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      stylePct.setDataFormat((short)0xa);
      
      if (!m_by_Dept) {
         m_CellStyles.add(styleText);    // col 0 date_break 
         m_CellStyles.add(stylePct);    // col 1 req_metric
         m_CellStyles.add(stylePct);    // col 2 allo_metric
         m_CellStyles.add(stylePct);    // col 3 ship_metric
         m_CellStyles.add(stylePct);    // col 4 perfect_metric
         if (m_Include_Totals){
        	if (m_Lines_Units_Dollars.equals("Dollars")){ 
               m_CellStyles.add(styleMoney);    // col 5 #req
               m_CellStyles.add(styleMoney);    // col 6 #plan
               m_CellStyles.add(styleMoney);    // col 7 #ship
               m_CellStyles.add(styleMoney);    // col 8 #Perf
               m_CellStyles.add(styleMoney);    // col 9 #total
        	}
        	else{ // lines or units
               m_CellStyles.add(styleInt);    // col 5 #req
               m_CellStyles.add(styleInt);    // col 6 #plan
               m_CellStyles.add(styleInt);    // col 7 #ship
               m_CellStyles.add(styleInt);    // col 8 #Perf
               m_CellStyles.add(styleInt);    // col 9 #total
       		
        	}
         }   
      }
      else {
         m_CellStyles.add(styleText);   // col 0 date_break
         m_CellStyles.add(styleText);   // col 1 dept          
         m_CellStyles.add(stylePct);    // col 2 req_metric
         m_CellStyles.add(stylePct);    // col 3 allo_metric
         m_CellStyles.add(stylePct);    // col 4 ship_metric
         m_CellStyles.add(stylePct);    // col 5 perfect_metric
         if (m_Include_Totals){
            if (m_Lines_Units_Dollars.equals("Dollars")){ 
        	   m_CellStyles.add(styleMoney);    // col 6 #req
               m_CellStyles.add(styleMoney);    // col 7 #plan
               m_CellStyles.add(styleMoney);    // col 8 #ship
               m_CellStyles.add(styleMoney);    // col 9 #Perf
               m_CellStyles.add(styleMoney);    // col 10 #total
            }
            else { // lines or units
         	   m_CellStyles.add(styleInt);    // col 6 #req
               m_CellStyles.add(styleInt);    // col 7 #plan
               m_CellStyles.add(styleInt);    // col 8 #ship
               m_CellStyles.add(styleInt);    // col 9 #Perf
               m_CellStyles.add(styleInt);    // col 10 #total
            }
         }      	  
      }
      styleText = null;
      styleInt = null;
      styleMoney = null;
      stylePct = null;
   }
}
