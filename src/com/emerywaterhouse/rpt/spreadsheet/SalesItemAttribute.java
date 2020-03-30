/**
 * File: SalesItemAttribute.java
 * Description: Report of Customer Sales by item attribute(s).  This report was originally
 *    created on 07/06/2005 by Jacob Heric;  It's been converted to the new report server
 *    format. 
 *
 * @author Jacob Heric
 * @author Jeffrey Fisher
 *
 * Create Date: 08/03/2005
 * Last Update: $Id: SalesItemAttribute.java,v 1.13 2009/02/18 16:53:10 jfisher Exp $
 * 
 * History
 *    $Log: SalesItemAttribute.java,v $
 *    Revision 1.13  2009/02/18 16:53:10  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.12  2008/10/29 20:55:02  jfisher
 *    Fixed potential null warnings.
 *
 *    Revision 1.11  2006/02/23 16:02:39  jfisher
 *    removed reference to logger and used the static logger in the report object.
 *
 *    Revision 1.10  2005/10/14 12:03:59  jheric
 *    Order the attribute name query (ensures column alignment is correct).
 *
 *    Revision 1.9  2005/10/12 13:54:25  jheric
 *    Display all Item Attributes (no matter how many).
 *
 *    08/24/2005 - Rewrite SQL because attribute/value conditions weren't working properly. JBH
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

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


public class SalesItemAttribute extends Report
{
   private String m_Attrs;
   private String m_AttrVals;
   private String m_BegDate;
   private String m_EndDate;
   private String m_FLC;   
   private int m_NumCols = 18;
   private int m_AttrCols = 7;  //This represents the number of columns in the report for
                                  //item attributes.   
   private String m_Sum;
   private boolean m_SumByCust; // sum by cust type reports are (very slightly) different
   
   private PreparedStatement m_SalesByAttr;
   private PreparedStatement m_ItemAttr;
   private PreparedStatement m_AttrValue;    
   private CallableStatement m_BaseSellPrice;
   private CallableStatement m_AMarket; 
   private CallableStatement m_BMarket;
   private CallableStatement m_CMarket;
   private CallableStatement m_DMarket;
   private CallableStatement m_EmeryCost;
      
   /**
    * default constructor
    */
   public SalesItemAttribute()
   {
      super();
      
      StringBuffer tmp = null;
      SimpleDateFormat dtf = null;
      
      try {
         //
         // Multiple instances of this report could be run simultaneously, 
         // give report unique name.
         tmp = new StringBuffer();
         dtf = new SimpleDateFormat("yyMMddHHmmss");
         tmp.append("salesitmattr");
         tmp.append(dtf.format(System.currentTimeMillis()));
         tmp.append(".xlsx");
         
         m_FileNames.add(tmp.toString());
      }
      
      finally {
         tmp = null;
         dtf = null;
      }
   }

   /**
    * Executes the queries and builds the output file
    */
   private boolean buildOutputFile()
   {      
      XSSFWorkbook WrkBk = null;
      XSSFSheet Sheet = null;
      XSSFRow Row = null;      
      FileOutputStream OutFile = null;
      //short ColCnt = 1;
      int RowNum = 1;
      boolean Result = false;
      //String[] list;      
      StringBuffer tmp = new StringBuffer();
      ResultSet item_attrs = null;
      ResultSet sales = null;
      int col;   
      double margin;         // Margin in Dollars
      double marginPerc;     // Margin in Percent
      double avgOrdCnt;      // Average Order Count
      double cost;           // Emery Cost
      double base;           // Emery Base
      double aMarket;        // A Market Retail
      double bMarket;        // B Market Retail
      double cMarket;        // C Market Retail
      double dMarket;        // D Market Retail 
      double totBuy;         // Total Buy
      double totSell;        // Total Sell      
      String item;           // item identifier
      int totLines;          // total lines shipped
      int totQtyShipped;     // total quantity shipped
      SimpleDateFormat fmt = new SimpleDateFormat("MM/dd/yyyy");
      
      try{
         OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         WrkBk = new XSSFWorkbook();
         Sheet = WrkBk.createSheet();
         
         //
         // Build progress message
         tmp.setLength(0);
         tmp.append("Retrieving data for report");
         setCurAction(tmp.toString());          
         
         m_SalesByAttr.setDate(1, new java.sql.Date(fmt.parse(m_BegDate).getTime()));
         m_SalesByAttr.setDate(2, new java.sql.Date(fmt.parse(m_EndDate).getTime()));         
         sales = m_SalesByAttr.executeQuery();
         
         //
         // Build progress message
         tmp.setLength(0);
         tmp.append("Building file for report");         
         setCurAction(tmp.toString()); 
         
         createCaptions(Sheet);
                  
         while ( sales.next() && m_Status == RptServer.RUNNING ) {            
            margin = 0.0;         // Margin in Dollars
            marginPerc = 0.0;     // Margin in Percent
            cost = 0.0;           // Emery Cost
            base = 0.0;           // Emery Base
            aMarket = 0.0;       // A Market Retail
            bMarket = 0.0;       // B Market Retail
            cMarket = 0.0;       // C Market Retail
            dMarket = 0.0;       // D Market Retail             
            
            //
            // Build progress message
            tmp.setLength(0);
            tmp.append("Adding item (");
            tmp.append(sales.getString("item_id"));
            tmp.append(") sales data to file for report");
            setCurAction(tmp.toString());            
            
            Row = createRow(Sheet, RowNum);
            RowNum++;
            col = -1;
            
            //
            //Store the item id for easy access
            item = sales.getString("item_id");
            
            // Get retail, cost & base
            cost = getEmeryCost(item) ;
            base = getBaseSellPrice(item);
            aMarket = getAMarketRet(item);
            bMarket = getBMarketRet(item);
            cMarket = getCMarketRet(item);
            dMarket = getDMarketRet(item);
            
            //
            //Store total buy & sell for easy access 
            totBuy = sales.getDouble("total_cost");
            totSell = sales.getDouble("total_sell");

            //
            //Dollar Margin
            margin = totSell - totBuy;
            margin = Math.floor(margin * 100 + .5d) / 100;

            //
            // Percentage Margin (prevent divide by zero)
            if ( totSell > 0.0 ) {
               marginPerc = (margin/totSell) * 100;
               marginPerc = Math.floor(marginPerc * 100 + .5d) / 100;
            } 
            
            //
            // Calculate average order quantity
            totLines = sales.getInt("lines");
            totQtyShipped= sales.getInt("qty_shipped");
            avgOrdCnt = Math.floor( (double)totQtyShipped/(double)totLines * 10 + .5 ) / 10;
            
            //
            //At long last, insert data
            if (m_SumByCust)
               Row.getCell(++col).setCellValue(new XSSFRichTextString(sales.getString("cust_type"))); // Customer Type
            
            Row.getCell(++col).setCellValue(new XSSFRichTextString(item)); // item#
            
            //
            //This report now shows all attribute values for every item (note 
            //that this sql employs a bit of subselectery and outter joinery to
            //ensure we get a row for every attribute whether or not the item
            //has a value for this attribute, this ensures the values end up in the
            //correct columns).
            //Added 10/11/2005 at M. Smith's request, after much squabbling. jbh
            m_ItemAttr.setString(1, item);
            item_attrs = m_ItemAttr.executeQuery();
            
            //
            //Each row gets a column, regardless of value (see comment immediately previous)
            while(item_attrs.next()){
               Row.getCell(++col).setCellValue(new XSSFRichTextString(item_attrs.getString("value")));   
            }
            
            Row.getCell(++col).setCellValue(new XSSFRichTextString(sales.getString("name")));
            Row.getCell(++col).setCellValue(new XSSFRichTextString(sales.getString("upc_code")));
            Row.getCell(++col).setCellValue(new XSSFRichTextString(sales.getString("description")));
            Row.getCell(++col).setCellValue(new XSSFRichTextString(sales.getString("flc_id")));
            Row.getCell(++col).setCellValue(sales.getInt("stock_pack"));
            
            //
            // Do a little NBC translating 
            if (sales.getString("nbc").equals("ALLOW BROKEN CASES"))
               Row.getCell(++col).setCellValue(new XSSFRichTextString(""));
            else 
               Row.getCell(++col).setCellValue(new XSSFRichTextString("N"));
            
            Row.getCell(++col).setCellValue(totQtyShipped); // units sold 
            Row.getCell(++col).setCellValue(totSell); // total/extended sales
            Row.getCell(++col).setCellValue(marginPerc); // emery GM%
            Row.getCell(++col).setCellValue(margin); // emery GM$
            Row.getCell(++col).setCellValue(avgOrdCnt); // average order qty
            Row.getCell(++col).setCellValue(cost); // cost
            Row.getCell(++col).setCellValue(base); // base
            Row.getCell(++col).setCellValue(aMarket); // A Market
            Row.getCell(++col).setCellValue(bMarket); // B Market
            Row.getCell(++col).setCellValue(cMarket); // C Market
            Row.getCell(++col).setCellValue(dMarket); // D Market
         }
         
         WrkBk.write(OutFile);
         WrkBk.close();
         Result = true;

      }
      catch ( Exception ex ) {
         Result = false;
         
         log.error("exception", ex);         
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         WrkBk = null;
         tmp  = null;

         if ( OutFile != null ) {
            try {
               OutFile.close();
            }
   
            catch( Exception e ) {
               ;
            }
            
            OutFile = null;
         }
         
         if (sales != null ) {
            try {
               sales.close();
            }
            catch ( SQLException e) {
            }
            sales = null;
         }
         
         if (item_attrs != null ) {
            try {
               item_attrs.close();
            }
            catch ( SQLException e) {
            }
            item_attrs = null;
         }                
      }         
      
      return Result;
      
   }
   
   /**
    * Builds the SQL to run the report
    *
    */
   private String buildSQL() throws Exception{
      String[] list;
      StringBuffer sql = new StringBuffer();
      StringBuffer attrs = new StringBuffer();
      StringBuffer values = new StringBuffer();
      StringBuffer types = new StringBuffer();
      boolean attrs_cond = false;      
      boolean values_cond = false;      
      boolean types_cond = false;
      ResultSet tmp = null;
      
      //
      //Generate attributes condition
      list = parseParam(m_Attrs);
      attrs_cond = list != null && list.length > 0;
      
      if ( attrs_cond ) {
         attrs.append(" and (");
         
         if ( list != null ) {
            for (int i = 0; i < list.length; i++){
               if (i == 0 ){
                  attrs.append("item.item_Id in ( select item_id from item_attribute, ");
                  attrs.append("attribute_value, attribute where ");
                  attrs.append("item_attribute.attribute_value_id = attribute_value.attribute_value_id ");
                  attrs.append("and attribute_value.attribute_id = attribute.attribute_id and attribute.name = '");
               }
               else {
                  attrs.append("and item.item_Id in ( select item_id from item_attribute, ");
                  attrs.append("attribute_value, attribute where ");
                  attrs.append("item_attribute.attribute_value_id = attribute_value.attribute_value_id ");
                  attrs.append("and attribute_value.attribute_id = attribute.attribute_id and attribute.name = '");
               }               
               
               attrs.append(list[i]);
               attrs.append("')");
            }
         }
         
         attrs.append(")");
      }
      
      //
      //Generate attribute values condition.
      //Note:  This list contains both attribute and values, this is necessary
      // because attribute values are only unique within a given attribute.
      list = parseParam(m_AttrVals);
      values_cond = list != null && list.length > 0;
      
      if ( values_cond ) {
         values.append(" and (");
         
         if ( list != null ) {
            for (int i = 0; i < list.length; i++){
               if (i == 0 ){
                  values.append("item.item_id in ( select item_id from item_attribute where ");
                  values.append("item_attribute.attribute_value_id = ");
               }
               else{ 
                  values.append("and item.item_id in ( select item_id from item_attribute where ");
                  values.append("item_attribute.attribute_value_id = ");
               }
   
               //
               //parse out attribute name and value, get value_id
               m_AttrValue.setString(1, list[i].substring(0, list[i].indexOf(",")).trim());
               m_AttrValue.setString(2, list[i].substring(list[i].indexOf(",") + 1).trim());         
               tmp = m_AttrValue.executeQuery();
               
               try {
                  if ( tmp.next() )
                     values.append(tmp.getInt("attribute_value_id"));
                  
                  values.append(")");
               }
                              
               finally {
                  DbUtils.closeDbConn(null, null, tmp);
                  tmp = null;         
               }
            }
         }
         
         values.append(")");
      }
      
      //
      //Generate customer types condition
      list = parseParam(m_Sum);
      types_cond = list != null && list.length > 0;
      
      //
      //If we are going to be building a sum by cust type report, note that in a
      //member variable because if affects the report format.
      m_SumByCust = types_cond;
      
      if ( types_cond ) {
         types.append(" and (");
         
         if ( list != null ) {
            for (int i = 0; i < list.length; i++){
               if (i == 0 )
                  types.append("market_class.description = '");
               else 
                  types.append(" or market_class.description = '");
               
               types.append(list[i]);
               types.append("'");
            }
         }
         
         types.append(")");
      }
      
      //
      // If their are customer type criteria, build appropriate sql:
      if ( types_cond ) {
         //
         // Get item sales information 
         sql.append("select market_class.description as cust_type, item.item_id, item.description, ");
         sql.append("item.flc_id, item.stock_pack, broken_case.description as nbc, upc.upc_code, vendor.name, ");
         sql.append("sum(inv_dtl.qty_shipped) as qty_shipped, sum(inv_dtl.unit_sell) as total_sell, ");
         sql.append("sum(unit_cost) as total_cost, count(inv_dtl.item_nbr) as lines ");
         sql.append("from inv_dtl, item, broken_case, vendor, cust_market, market_class, ");
         sql.append("(select item_id, upc_code from item_upc where primary_upc = 1) upc ");
         sql.append("where inv_dtl.item_nbr = item.item_id and ");
         sql.append("item.vendor_id = vendor.vendor_id and ");         
         sql.append("item.item_id = upc.item_id(+) and ");
         sql.append("inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("item.broken_case_id = broken_case.broken_case_id and ");
         sql.append("inv_dtl.invoice_date >= ? and ");
         sql.append("inv_dtl.invoice_date <= ? and ");
         sql.append("inv_dtl.cust_nbr = cust_market.customer_id and "); 
         sql.append("cust_market.mkt_class_id = market_class.mkt_class_id ");
         sql.append(types);
         
         if (attrs_cond)
            sql.append(attrs);
         
         if (values_cond)
            sql.append(values);
         
         //
         //Generate the flc condition
         if (m_FLC != null && m_FLC.length() > 0) {
            sql.append(" and item.flc_id = '");
            sql.append(m_FLC);
            sql.append("'");
         }         

         sql.append(" group by market_class.description, item.item_id, item.description, ");
         sql.append("item.flc_id, item.stock_pack, broken_case.description, upc.upc_code, vendor.name ");      
      
      }
      else {
         //
         // Get item sales information 
         sql.append("select item.item_id, item.description, item.flc_id, item.stock_pack, ");
         sql.append("broken_case.description as nbc, upc.upc_code, vendor.name, sum(inv_dtl.qty_shipped) as qty_shipped, ");
         sql.append("sum(inv_dtl.unit_sell) as total_sell, sum(unit_cost) as total_cost, ");      
         sql.append("count(inv_dtl.item_nbr) as lines ");
         sql.append("from inv_dtl, item, broken_case, vendor, ");
         sql.append("(select item_id, upc_code from item_upc where primary_upc = 1) upc ");
         sql.append("where inv_dtl.item_nbr = item.item_id and ");
         sql.append("item.vendor_id = vendor.vendor_id and ");         
         sql.append("item.item_id = upc.item_id(+) and ");
         sql.append("inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("item.broken_case_id = broken_case.broken_case_id and ");
         sql.append("inv_dtl.invoice_date >= ? and ");
         sql.append("inv_dtl.invoice_date <= ? ");
         
         if ( attrs_cond )
            sql.append(attrs);
         
         if ( values_cond )
            sql.append(values);
         
         //
         //Generate the flc condition
         if (m_FLC != null && m_FLC.length() > 0) {
            sql.append(" and item.flc_id = '");
            sql.append(m_FLC);
            sql.append("'");
         }              
         
         sql.append(" group by item.item_id, item.description, item.flc_id, item.stock_pack, ");
         sql.append("broken_case.description, upc.upc_code, vendor.name ");         
      }
             
      return sql.toString();
   }
    
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      if ( m_SalesByAttr != null ) {
         try {
            m_SalesByAttr.close();
         }
         catch ( SQLException e) {
         }
         m_SalesByAttr = null;
      }
      
      if ( m_AttrValue != null ) {
         try {
            m_AttrValue.close();
         }
         catch ( SQLException e) {
         }
         m_AttrValue = null;
      }      
      
      if ( m_BaseSellPrice != null ) {
         try {
            m_BaseSellPrice.close();
         }
         catch ( SQLException e) {
         }
         m_BaseSellPrice = null;
      }

      if ( m_AMarket != null ) {
         try {
            m_AMarket.close();
         }
         catch ( SQLException e) {
         }
         m_AMarket = null;
      }
      
      if ( m_BMarket != null ) {
         try {
            m_BMarket.close();
         }
         catch ( SQLException e) {
         }
         m_BMarket = null;
      }
      
      if ( m_CMarket != null ) {
         try {
            m_CMarket.close();
         }
         catch ( SQLException e) {
         }
         m_CMarket = null;
      }
      
      if ( m_DMarket != null ) {
         try {
            m_DMarket.close();
         }
         catch ( SQLException e) {
         }
         m_DMarket = null;
      }
      
      if ( m_EmeryCost != null ) {
         try {
            m_EmeryCost.close();
         }
         catch ( SQLException e) {
         }
         m_EmeryCost = null;
      }      
      
      if ( m_ItemAttr != null ) {
         try {
            m_ItemAttr.close();
         }
         catch ( SQLException e) {
         }
         m_EmeryCost = null;
      }      
   }
   
   /**
    * Builds the captions on the worksheet.
    * Caption list is now highly variable.  It must encompass all possible
    * Item Attributes (see table eis_emery.attribute).  Added 10/11/2005. jbh
    */
   private void createCaptions(XSSFSheet sheet) throws Exception
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;
      int col = -1;
      StringBuffer tmp = new StringBuffer();
      ResultSet attr_names = null;  
      Statement attr_sql = null; 

      if ( sheet == null )
         return;

      Row = sheet.createRow(0);
      
      if ( m_SumByCust )
         ++m_NumCols;
      
      tmp = new StringBuffer();
      tmp.append("select name from attribute order by name");
      
      try {
         //
         //Get the list of all possible attributes
         attr_sql = m_EdbConn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,
               ResultSet.CONCUR_READ_ONLY);
         attr_names = attr_sql.executeQuery(tmp.toString());
         //
         //We need to know the number of attributes 
         //before we step through them (so we can 
         //create a row of sufficient length).
         attr_names.last(); 
         m_AttrCols =  attr_names.getRow();
         m_NumCols =  (m_NumCols + m_AttrCols);
         attr_names.beforeFirst();
         
         if ( Row != null ) {
            for ( int i = 0; i < m_NumCols; i++ ) {
               Cell = Row.createCell(i);
               Cell.setCellType(CellType.STRING);
            }
   
            if ( m_SumByCust )
               Row.getCell(++col).setCellValue(new XSSFRichTextString("Customer Type"));
            
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Item"));
            
            //
            //Add all possible attributes to report caption
            while (attr_names.next()){
               m_NumCols++;
               Row.getCell(++col).setCellValue(new XSSFRichTextString(attr_names.getString("name")));
            }              
            
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Vendor"));         
            Row.getCell(++col).setCellValue(new XSSFRichTextString("UPC"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Description"));          
            Row.getCell(++col).setCellValue(new XSSFRichTextString("FLC"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Stock Pack"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("NBC"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Unit Sales"));         
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Sales"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("GM(%)"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("GM($)"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Avg. Order Qty."));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Cost"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("Base"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("A Market"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("B Market"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("C Market"));
            Row.getCell(++col).setCellValue(new XSSFRichTextString("D Market"));
         }
      
      }
      catch (Exception e){
         throw new Exception(e.getMessage());
      }
      finally{
         
         if ( attr_names != null ) {
            try {
               attr_names.close();
               attr_names = null;
            }
            catch ( Exception e ) {
            }
         }
         
         if ( attr_sql != null ) {
            try {
               attr_sql.close();
               attr_sql = null;
            }
            catch ( Exception e ) {
            }
         }

         tmp = null;         
         
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
         m_EdbConn = m_RptProc.getEdbConn();
         prepareStatements();
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
    * Creates a row in the worksheet.
    */
   private XSSFRow createRow(XSSFSheet sheet, int rowNum)
   {
      XSSFRow row = sheet.createRow(rowNum);
      XSSFCell cell;
      int col = -1;

      //
      // All rows will have this height
      //row.setHeightInPoints(ROW_HEIGHT);

      //
      // Add the cells
      if (m_SumByCust){
         cell = row.createCell(++col);  // Customer Type
         cell.setCellType(CellType.STRING);
      }
      
      cell = row.createCell(++col);  // Item
      cell.setCellType(CellType.STRING);
      
      //
      //Add however many attribute columns there are
      for (int i = 1; i <= m_AttrCols; i++ ){
         cell = row.createCell(++col);  // Item Attribute
         cell.setCellType(CellType.STRING);
      }
         
      
      //
      //Added 08/30/2005 at the request of mark smith. jbh
      cell = row.createCell(++col);  // Vendor Name
      cell.setCellType(CellType.STRING);      

      cell = row.createCell(++col);  // UPC
      cell.setCellType(CellType.STRING);

      cell = row.createCell(++col);  // Description
      cell.setCellType(CellType.STRING);

      cell = row.createCell(++col);  // FLC
      cell.setCellType(CellType.STRING);

      cell = row.createCell(++col);  // STOCK PACK
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // NBC
      cell.setCellType(CellType.STRING);

      cell = row.createCell(++col);  // Unit Sales
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Sales
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // GM%
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // GM$
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Average Order Qty.
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cost
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Base
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // A Market
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // B Market
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // C Market
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // D Market
      cell.setCellType(CellType.NUMERIC);

      cell = null;

      return row;
   }
   
   /**
    * Gets the current a market retail, on exception, returns zero
    *
    * @param itemId String - the input item identifier.
    * @return double - a market retail
    */
   public double getAMarketRet(String itemId)
   {
      double ret = 0.0;

      try {
         m_AMarket.setString(2, itemId);
         m_AMarket.execute();

         ret = m_AMarket.getDouble(1);
      }
      
      catch ( SQLException e1 ) {
         ret = 0.0;
      }

      return ret;
   }
   
   /**
    * Gets the current b market retail, on exception, returns zero
    *
    * @param itemId String - the input item identifier.
    * @return double - b market retail
    */
   public double getBMarketRet(String itemId)
   {
      double ret = 0.0;

      try {
         m_BMarket.setString(2, itemId);
         m_BMarket.execute();

         ret = m_BMarket.getDouble(1);
      }
      
      catch ( SQLException e1 ) {
         ret = 0.0;
      }

      return ret;
   }         
   
   /**
    * Gets the current c market retail, on exception, returns zero
    *
    * @param itemId String - the input item identifier.
    * @return double - c market retail
    */
   public double getCMarketRet(String itemId)
   {
      double ret = 0.0;

      try {
         m_CMarket.setString(2, itemId);
         m_CMarket.execute();

         ret = m_CMarket.getDouble(1);
      }
      
      catch ( SQLException e1 ) {
         ret = 0.0;
      }

      return ret;
   }         
   
   /**
    * Gets the current d market retail, on exception, returns zero
    *
    * @param itemId String - the input item identifier.
    * @return double - d market retail
    */
   public double getDMarketRet(String itemId)
   {
      double ret = 0.0;

      try {
         m_DMarket.setString(2, itemId);
         m_DMarket.execute();

         ret = m_DMarket.getDouble(1);
      }
      
      catch ( SQLException e1 ) {
         ret = 0.0;
      }

      return ret;
   }         
   
   /**
    * Gets the current base sell price, on exception, returns zero
    *
    * @param itemId String - the input item identifier.
    * @return double - the current base sell price.
    */
   public double getBaseSellPrice(String itemId)
   {
      double base = 0.0;

      try {
         m_BaseSellPrice.setString(2, itemId);
         m_BaseSellPrice.execute();

         base = m_BaseSellPrice.getDouble(1);
      }
      
      catch ( SQLException e1 ) {
         base = 0.0;
      }

      return base;
   }
   
   /**
    * Gets the current emery cost, on exception, returns zero
    *
    * @param itemId String - the input item identifier.
    * @return double - the current emery cost
    */
   public double getEmeryCost(String itemId)
   {
      double cost = 0.0;

      try {
         m_EmeryCost.setString(2, itemId);
         m_EmeryCost.execute();

         cost = m_EmeryCost.getDouble(1);
      }
      
      catch ( SQLException e1 ) {
         cost = 0.0;
      }

      return cost;
   }      

   /**
    * Parses a semicolon separated list of parameters
    *
    * @param paramList String - the list of params to be parsed.
    * @return String[] An array of parameter values to use in this report
    */
   public String[] parseParam(String paramList)
   {
      String list[] = null;
      int i = 0;
      int j = 0;
      int len = 0;
      ArrayList<String> tmpList = new ArrayList<String>();
      
      //
      // We have to create a temporary holding place for the values since we don't know
      // how many we are going to have.
      if ( paramList != null && paramList.length() > 0) {
         while ( i != -1 ) {
            i = paramList.indexOf(';', i);
   
            if ( i != -1 ) {       
               tmpList.add(paramList.substring(j, i));
               
               i++;
               j = i;
            }
            else {
               //
               // Handle the case of one value or the last one in the list.
               i = paramList.length();
   
               if ( j < i ) {            
                  tmpList.add(paramList.substring(j, i));
                  i = -1;
               }
            }
         }
      }

      //
      // Once we have the list of values we can get a count and create the array.
      len = tmpList.size();
      j = 0;

      if ( len > 0 ) {
         list = new String[len];

         for ( i = 0; i < len; i++ ) {
            list[i] = tmpList.get(i);
         }
      }

      return list;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private void prepareStatements() throws Exception
   {  
      StringBuffer tmp = new StringBuffer();
      if ( m_EdbConn != null ) {
         
         //
         //attribute_value_id query, this must get prepared before m_SalesByAttr because 
         //buildSQL uses it.
         tmp.setLength(0);
         tmp.append("select attribute_value_id ");
         tmp.append("from attribute, attribute_value "); 
         tmp.append("where attribute.attribute_id = attribute_value.attribute_id and ");
         tmp.append("attribute.name = ? and attribute_value.value = ? ");
         m_AttrValue = m_EdbConn.prepareStatement(tmp.toString());
         
         //
         //Query to get a list of all possible attributes (and their values) for 
         //an item (even if then there is no value).
         tmp.setLength(0);
         tmp.append("select name, av.value ");
         tmp.append("from attribute, ");
         tmp.append("(select value, attribute_id from attribute_value, item_attribute ");
         tmp.append("where attribute_value.attribute_value_id = item_attribute.attribute_value_id and ");
         tmp.append("item_attribute.item_id = ?) av ");
         tmp.append("where attribute.attribute_id = av.attribute_id(+) " );
         tmp.append("order by attribute.name " );          
         m_ItemAttr = m_EdbConn.prepareStatement(tmp.toString());
         
         //
         //Main report query, the complicates sql get built in it's own method
         m_SalesByAttr = m_EdbConn.prepareStatement(buildSQL());

         //
         // Gets base sell price.  
         m_BaseSellPrice = m_EdbConn.prepareCall(
            "call item_price_procs.todays_sell(?)"
         );
         
         //
         // Gets a market retail.  
         m_AMarket = m_EdbConn.prepareCall(
            "call item_price_procs.todays_retaila(?)"
         );
         
         //
         // Gets the b market retail.  
         m_BMarket = m_EdbConn.prepareCall(
            "call item_price_procs.todays_retailb(?)"
         );
         
         //
         // Gets c market retail.  
         m_CMarket = m_EdbConn.prepareCall(
            "call  item_price_procs.todays_retailc(?)"
         );
         
         //
         // Gets the d market retail.  
         m_DMarket = m_EdbConn.prepareCall(
            "call  item_price_procs.todays_retaild(?)"
         );
         
         //
         // Gets the current emery cost for the input item.
         m_EmeryCost = m_EdbConn.prepareCall(
            "call  item_price_procs.todays_buy(?)"
         );
      }      
   }
   
   /**
    * Sets the parameters for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      m_BegDate = params.get(0).value;
      m_EndDate = params.get(1).value;
      m_Attrs = params.get(2).value;
      m_AttrVals = params.get(3).value;
      m_Sum = params.get(4).value;
      m_FLC = params.get(5).value;      
   }

}
