/**
 * 
 * File: NewDiscItem.java
 * Description: The New item/Discontinued item/New Catalog item report excel spreadsheet.
 * Company:     Emery-Waterhouse
 * @author      bcornwell
 * Create Date: 03/20/2018
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;


public class NewDiscItem extends Report 
{
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row;
   private XSSFFont m_FontNorm;
   private XSSFFont m_FontBold;
   private XSSFCellStyle m_StyleHdrLeft;
   private XSSFCellStyle m_StyleCaption;
   private XSSFCellStyle m_StyleTxtC;      // Text centered
   private XSSFCellStyle m_StyleTxtL;      // Text left justified
   private XSSFCellStyle m_StyleInt;       // Style with 0 decimals
   private XSSFCellStyle m_StyleDouble;    // numeric #,##0.00
   
   private PreparedStatement m_RptData;
   private int m_Warehouse;
   private String m_BegDate;
   private String m_EndDate;
   private int m_RptType;
   private static final int NEW_ITEM_REPORT = 0;
   private static final int DISCONTINUED_ITEM_REPORT = 1;
   private static final int CATALOG_ITEM_REPORT = 2;
   private static final int PORTLAND_REQUESTED = 1;
   private static final int PITTSTON_REQUESTED = 2;
   
   /**
    * default constructor
    */
   public NewDiscItem()
   {
   	super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      
	   defineStyles();
   }
   
   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {            
      m_Sheet = null;
      m_Wrkbk = null;
      
      m_StyleHdrLeft = null;
      m_StyleCaption = null;
      m_StyleTxtC = null;
      m_StyleTxtL = null;
      m_StyleInt = null;
      m_StyleDouble = null;
      
      super.finalize();
   }
   
   /**
    * adds a numeric type cell to current row at column col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value numeric value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, double value, XSSFCellStyle style)
   {
      XSSFCell cell = m_Row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(value);
      
      cell = null;
   }
   
   /**
    * adds a text type cell to current row at column col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value text value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, String value, XSSFCellStyle style)
   {
	   XSSFCell cell = m_Row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(value));
      cell.setCellStyle(style);
      cell = null;
   }
   
   /**
    * adds row to the current sheet
    *
    * @param row  0-based row number of row to be added
    */
   private void addRow(int row) 
   {
      m_Row = m_Sheet.createRow(row);
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      boolean result = false;
      ResultSet rptData = null;
      int colNum = 0;
      int rowNum = 0;
      
      try {
          switch (m_RptType) {
          	case  NEW_ITEM_REPORT:
          		rowNum = createNewItemCaptions();
          		m_RptData.setString(1, m_BegDate);
          		m_RptData.setString(2, m_EndDate);
          		m_RptData.setString(3, m_BegDate);
          		m_RptData.setString(4, m_EndDate);
          		m_RptData.setString(5, m_BegDate);
          		m_RptData.setString(6, m_EndDate);
          		break;
          	case DISCONTINUED_ITEM_REPORT:
          		rowNum = createDiscontinuedItemCaptions();
          		m_RptData.setString(1, m_BegDate);
          		m_RptData.setString(2, m_EndDate);
          		break;
 	         case CATALOG_ITEM_REPORT:
 	         	rowNum = createCatalogItemCaptions();
 	         	m_RptData.setString(1, m_BegDate);
 	         	m_RptData.setString(2, m_EndDate);
 	         	break;
          }
         
          rptData = m_RptData.executeQuery();
         
          while ( rptData.next() && m_Status == RptServer.RUNNING ) {
         	 addRow(rowNum++);
         	 colNum = 0;
         	 switch (m_RptType) {
         	 	case  NEW_ITEM_REPORT:
         	 		addCell(colNum++, rptData.getString("item_dept_num"), m_StyleTxtC);
         	 		addCell(colNum++, rptData.getString("item_vendor_name"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("item_id"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("item_nrha_id"), m_StyleTxtC);
         	 		addCell(colNum++, rptData.getString("item_flc_id"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("item_descr"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getDouble("item_sell"), m_StyleDouble);
         	 		addCell(colNum++, rptData.getDouble("item_retail_c"), m_StyleDouble);
         	 		addCell(colNum++, rptData.getString("item_upc"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("item_shipunit"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("item_stock_pack"), m_StyleInt);
         	 		addCell(colNum++, rptData.getString("item_nbc"), m_StyleTxtC);
         	 		addCell(colNum++, rptData.getString("soq_comment"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("ptldcat"), m_StyleTxtC);
         	 		addCell(colNum++, rptData.getString("pittcat"), m_StyleTxtC);
         	 		addCell(colNum++, rptData.getString("item_setup_date"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("statusdt"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("active_begin_date"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("actual_qty"), m_StyleInt);
         	 		addCell(colNum++, rptData.getString("item_disp"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getString("item_vendor_id"), m_StyleTxtL);
         	 		addCell(colNum++, rptData.getDouble("item_buy"), m_StyleDouble);
         	 		addCell(colNum++, rptData.getString("whs_name"), m_StyleTxtL);
         	 		break;
	 	         case DISCONTINUED_ITEM_REPORT:
	 	         	addCell(colNum++, rptData.getString("item_dept_num"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_vendor_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_vendor_name"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_descr"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_nrha_id"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_flc_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_upc"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_nbc"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_stock_pack"), m_StyleInt);
	 	         	addCell(colNum++, rptData.getString("item_shipunit"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_rms_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_rms_description"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getDouble("item_buy"), m_StyleDouble);
	 	         	addCell(colNum++, rptData.getDouble("item_sell"), m_StyleDouble);
	 	         	addCell(colNum++, rptData.getDouble("item_retail_c"), m_StyleDouble);
	 	         	addCell(colNum++, rptData.getString("soq_comment"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("catalog_page"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("auto_sub"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("sub_item_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_dept_num"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("sub_nrha_id"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("sub_vendor_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_vendor_name"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_flc_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_description"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_upc"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_nbc"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("sub_stock_pack"), m_StyleInt);
	 	         	if (rptData.getString("sub_buy") != null)
	 	         		addCell(colNum++, rptData.getDouble("sub_buy"), m_StyleDouble);
	 	         	else
	 	         		addCell(colNum++, "", m_StyleTxtL);
	 	         	if (rptData.getString("sub_sell") != null)
	 	         		addCell(colNum++, rptData.getDouble("sub_sell"), m_StyleDouble);
	 	         	else
	 	         		addCell(colNum++, "", m_StyleTxtL);
	 	         	if (rptData.getString("sub_retail_c") != null)
	 	         		addCell(colNum++, rptData.getDouble("sub_retail_c"), m_StyleDouble);
	 	         	else
	 	         		addCell(colNum++, "", m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_shipunit"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_rms_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("sub_rms"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_disp"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("whs_name"), m_StyleTxtL);
	 	         	break;
	 	         case CATALOG_ITEM_REPORT:
	 	         	addCell(colNum++, rptData.getString("item_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_descr"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_dept_num"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_vendor_id"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_vendor_name"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_nrha_id"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_flc_id"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_upc"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_nbc"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("item_stock_pack"), m_StyleInt);
	 	         	addCell(colNum++, rptData.getString("item_shipunit"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("ptldcat"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getString("pittcat"), m_StyleTxtC);
	 	         	addCell(colNum++, rptData.getDouble("item_buy"), m_StyleDouble);
	 	         	addCell(colNum++, rptData.getDouble("item_sell"), m_StyleDouble);
	 	         	addCell(colNum++, rptData.getDouble("item_retail_c"), m_StyleDouble);
	 	         	addCell(colNum++, rptData.getString("soq_comment"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("item_setup_date"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("statusdt"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("actual_qty"), m_StyleInt);
	 	         	addCell(colNum++, rptData.getString("item_disp"), m_StyleTxtL);
	 	         	addCell(colNum++, rptData.getString("whs_name"), m_StyleTxtL);
	 	         	break;
         	 }
          }

          m_Wrkbk.write(outFile);
          rptData.close();
          result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[NewDiscItem]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, m_RptData, rptData);
      }
      
      return result;
   }
   
   /**
    * Creates the excel spreadsheet.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         if ( m_EdbConn == null && m_RptProc != null )
            m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();
         
         setCurAction("complete");
      }
      
      catch ( Exception ex ) {
         log.fatal("[NewDiscItem]", ex);
      }
      
      finally {                  
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;      
   }

   /**
    * defines Excel fonts and styles
    */
   private void defineStyles()
   {
      short fontHeight = 11;
      XSSFDataFormat format = m_Wrkbk.createDataFormat();
      
      //
      // Normal Font
      m_FontNorm = m_Wrkbk.createFont();
      m_FontNorm.setFontName("Arial");
      m_FontNorm.setFontHeightInPoints(fontHeight);
      
      //
      // defines bold font
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontName("Arial");
      m_FontBold.setFontHeightInPoints(fontHeight);
      m_FontBold.setBold(true);
      
      //
      // defines style column header, left-justified
      m_StyleHdrLeft = m_Wrkbk.createCellStyle();
      m_StyleHdrLeft.setFont(m_FontBold);
      m_StyleHdrLeft.setAlignment(HorizontalAlignment.LEFT);
      m_StyleHdrLeft.setVerticalAlignment(VerticalAlignment.TOP);

      m_StyleCaption = m_Wrkbk.createCellStyle();
      m_StyleCaption.setFont(m_FontBold);
      m_StyleCaption.setAlignment(HorizontalAlignment.CENTER);
      m_StyleCaption.setWrapText(true);
      
      m_StyleTxtL = m_Wrkbk.createCellStyle();
      m_StyleTxtL.setAlignment(HorizontalAlignment.LEFT);
      
      m_StyleTxtC = m_Wrkbk.createCellStyle();
      m_StyleTxtC.setAlignment(HorizontalAlignment.CENTER);
      
      m_StyleInt = m_Wrkbk.createCellStyle();
      m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleInt.setDataFormat((short)3);
      
      m_StyleDouble = m_Wrkbk.createCellStyle();
      m_StyleDouble.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDouble.setDataFormat(format.getFormat("$#,##0.00"));
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      String tmp = null;
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("startdate") ) 
           	m_BegDate = param.value;
         
         if ( param.name.equals("enddate") ) 
           	m_EndDate = param.value;
         
         if ( param.name.equals("warehouse") )
         	m_Warehouse = Integer.parseInt(param.value);

         if ( param.name.equals("rpttype") )
            m_RptType = Integer.parseInt(param.value);
      }
      
      tmp = Long.toString(System.currentTimeMillis());

      switch (m_RptType) {
      	case  NEW_ITEM_REPORT:
      		fileName.append("New_Items_");
      		break;
	      case DISCONTINUED_ITEM_REPORT:
	      	fileName.append("Discontinued_Items_");
	      	break;
	      case CATALOG_ITEM_REPORT:
	      	fileName.append("New_Catalog_Items_");
	      	break;
      }
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);      
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {
             sql.setLength(0);
             sql.append("select item_ed.dept_num as item_dept_num, ");
             sql.append("    warehouse.name as whs_name, ");
             sql.append("    item_iea.vendor_id as item_vendor_id, ");
             sql.append("    item_vendor.name as item_vendor_name,  ");
             sql.append("    item_iea.item_id, item_iea.description as item_descr, ");
             sql.append("    item_ei.flc_id as item_flc_id, ");
             sql.append("    to_char(item_ei.setup_date, 'mm/dd/yyyy') as item_setup_date, ");
             sql.append("    to_char(item_eiw.active_begin, 'mm/dd/yyyy') as active_begin_date, ");
             sql.append("    case when item_ei.broken_case_id = 1 then 'N' else 'Y' end as item_nbc, ");
             sql.append("    item_ei.soq_comment, ");
             sql.append("    item_mdc.nrha_id as item_nrha_id, ");
             sql.append("    item_su.name as item_shipunit, ");
             sql.append("    ( ");
             sql.append("      select upc_code ");
             sql.append("      from ejd_item_whs_upc ");
             sql.append("      where ejd_item_id = item_iea.ejd_item_id ");
             sql.append("        and ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id ");
             sql.append("      order by primary_upc desc, upc_code limit 1 ");
             sql.append("    ) as item_upc, ");
             sql.append("    ( ");
             sql.append("      select sum(loc_allocation.actual_qty) ");
             sql.append("      from loc_allocation ");
             sql.append("      where loc_allocation.sku = item_iea.item_id ");
             sql.append("        and loc_allocation.warehouse = warehouse.name ");
             sql.append("    ) as actual_qty, ");
             sql.append("    item_eiw.stock_pack as item_stock_pack, ");

             switch (m_RptType) {
	 	         case  NEW_ITEM_REPORT:
	 	         	sql.append("    (select to_char(max(statusdt.change_date), 'mm/dd/yyyy') ");
	 	         	sql.append("     from ejd_item_disp_history statusdt ");
	 	         	sql.append("     where statusdt.ejd_item_id = item_eiw.ejd_item_id ");
	 	         	sql.append("     and statusdt.warehouse_id = item_eiw.warehouse_id ");
	 	         	sql.append("     and statusdt.new_disp_id = 1 ");
	 	         	sql.append("    ) as statusdt, ");
	 	         	break;
	 	         case DISCONTINUED_ITEM_REPORT:
	 	         	sql.append("    (select to_char(max(statusdt.change_date), 'mm/dd/yyyy') ");
	 	         	sql.append("     from ejd_item_disp_history statusdt ");
	 	         	sql.append("     where statusdt.ejd_item_id = item_eiw.ejd_item_id ");
	 	         	sql.append("     and statusdt.warehouse_id = item_eiw.warehouse_id ");
	 	         	sql.append("     and (statusdt.new_disp_id = 2 or statusdt.new_disp_id = 3) ");
	 	         	sql.append("    ) as statusdt, ");
	 	         	sql.append("    web_item_ea.page as catalog_page, ");
	 	         	sql.append("    item_ri.rms_id as item_rms_id, item_rms.description as item_rms_description, ");
	 	         	sql.append("    case when item_ea_sub.auto_sub is null then '' when item_ea_sub.auto_sub = 1 then 'Y' else 'N' end as auto_sub, ");
	 	         	sql.append("    sub_ed.dept_num as sub_dept_num, ");
	 	         	sql.append("    sub_iea.item_id as sub_item_id, ");
	 	         	sql.append("    sub_iea.description as sub_description, ");
	 	         	sql.append("    sub_iea.vendor_id as sub_vendor_id, ");
	 	         	sql.append("    case when sub_ei.broken_case_id is null then '' when sub_ei.broken_case_id = 1 then 'N' else 'Y' end as sub_nbc, ");
	 	         	sql.append("    sub_vendor.name as sub_vendor_name,  ");
	 	         	sql.append("    sub_mdc.nrha_id as sub_nrha_id, ");
	 	         	sql.append("    sub_ei.flc_id as sub_flc_id, ");
	 	         	sql.append("    sub_su.name as sub_shipunit, ");
	 	         	sql.append("    sub_eiw.stock_pack as sub_stock_pack, ");
	 	         	sql.append("    sub_eip.buy as sub_buy, ");
	 	         	sql.append("    sub_eip.sell as sub_sell, ");
	 	         	sql.append("    sub_eip.retail_c as sub_retail_c, ");
	 	         	sql.append("    ( ");
	 	         	sql.append("      select upc_code ");
	 	         	sql.append("      from ejd_item_whs_upc ");
	 	         	sql.append("      where ejd_item_id = sub_iea.ejd_item_id ");
	 	         	sql.append("        and ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id ");
	 	         	sql.append("      order by primary_upc desc, upc_code limit 1 ");
	 	         	sql.append("     ");
	 	         	sql.append("    ) as sub_upc, ");
	 	         	sql.append("    sub_ri.rms_id as sub_rms_id, sub_rms.description as sub_rms, ");
	 	         	break;
	 	         case CATALOG_ITEM_REPORT:
	 	         	sql.append("    to_char(nci.new_cat_date ,'mm/dd/yyyy') as statusdt, ");
	 	         	break;
	 	       }
             
             sql.append("    item_id.disposition as item_disp, ");
             sql.append("    item_eip.buy as item_buy, ");
             sql.append("    item_eip.sell as item_sell, ");
             sql.append("    item_eip.retail_c as item_retail_c, ");
             sql.append("    case when ptld.in_catalog = 1 then 'Y' else 'N' end as ptldcat, ");
             sql.append("    case when pitt.in_catalog = 1 then 'Y' else 'N' end as pittcat, ");
             sql.append("    item_type.itemtype ");
             sql.append("from ejd_item item_ei ");
             sql.append("inner join item_entity_attr item_iea ");
             sql.append("  on item_iea.ejd_item_id = item_ei.ejd_item_id ");
             sql.append("inner join item_type ");
             sql.append("  on item_type.item_type_id = item_iea.item_type_id ");
             sql.append("inner join emery_dept item_ed ");
             sql.append("  on item_ed.dept_id = item_ei.dept_id ");
             sql.append("inner join vendor item_vendor ");
             sql.append("  on item_vendor.vendor_id = item_iea.vendor_id ");
             sql.append("inner join warehouse ");

             switch (m_Warehouse) {
	             case PORTLAND_REQUESTED:
	            	 sql.append("  on warehouse.warehouse_id = 1 ");
	            	 break;
	             case PITTSTON_REQUESTED:
	            	 sql.append("  on warehouse.warehouse_id = 2 ");
	            	 break;
	             default:
	            	 sql.append("  on warehouse.warehouse_id in (1, 2) ");
	            	 break;
             }
             sql.append("inner join ejd_item_warehouse item_eiw ");
             sql.append("  on item_eiw.ejd_item_id = item_ei.ejd_item_id ");
             sql.append("  and item_eiw.warehouse_id = warehouse.warehouse_id ");
             sql.append("inner join item_disp item_id ");
             sql.append("  on item_id.disp_id = item_eiw.disp_id ");
             sql.append("inner join ejd_item_price item_eip ");
             sql.append("  on item_eip.ejd_item_id = item_ei.ejd_item_id ");
             sql.append("  and item_eip.warehouse_id = warehouse.warehouse_id ");

             switch (m_RptType) {
	 	         case  NEW_ITEM_REPORT:
	 	             //
	 	         	break;
	 	         case DISCONTINUED_ITEM_REPORT:
	 	         	sql.append("left outer join web_item_ea ");
	 	         	sql.append("  on web_item_ea.item_ea_id = item_iea.item_ea_id ");
	 	         	sql.append("left outer join rms_item item_ri ");
	 	         	sql.append("  on item_ri.item_ea_id = item_iea.item_ea_id ");
	 	         	sql.append("left outer join rms item_rms  ");
	 	         	sql.append("  on item_rms.rms_id = item_ri.rms_id ");
	 	         	sql.append("left outer join item_ea_sub ");
	 	         	sql.append("  on item_ea_sub.item_ea_id = item_iea.item_ea_id ");
	 	         	sql.append("left outer join item_entity_attr sub_iea ");
	 	         	sql.append("  on sub_iea.item_ea_id = item_ea_sub.item_ea_id ");
	 	         	sql.append("left outer join ejd_item sub_ei ");
	 	         	sql.append("  on sub_ei.ejd_item_id = sub_iea.ejd_item_id ");
	 	         	sql.append("left outer join ejd_item_warehouse sub_eiw ");
	 	         	sql.append("  on sub_eiw.ejd_item_id = sub_ei.ejd_item_id ");
	 	         	sql.append("  and sub_eiw.warehouse_id = warehouse.warehouse_id ");
	 	         	sql.append("left outer join ejd_item_price sub_eip ");
	 	         	sql.append("  on sub_eip.ejd_item_id = sub_ei.ejd_item_id ");
	 	         	sql.append("  and sub_eip.warehouse_id = warehouse.warehouse_id ");
	 	         	sql.append("left outer join emery_dept sub_ed ");
	 	         	sql.append("  on sub_ed.dept_id = sub_ei.dept_id ");
	 	         	sql.append("left outer join vendor sub_vendor ");
	 	         	sql.append("  on sub_vendor.vendor_id = sub_iea.vendor_id ");
	 	         	sql.append("left outer join flc sub_flc ");
	 	         	sql.append("  on sub_flc.flc_id = sub_ei.flc_id ");
	 	         	sql.append("left outer join mdc sub_mdc ");
	 	         	sql.append("  on sub_mdc.mdc_id = sub_flc.mdc_id ");
	 	         	sql.append("left outer join ship_unit sub_su ");
	 	         	sql.append("  on sub_su.unit_id = sub_iea.ship_unit_id ");
	 	         	sql.append("left outer join rms_item sub_ri ");
	 	         	sql.append("  on sub_ri.item_ea_id = sub_iea.item_ea_id ");
	 	         	sql.append("left outer join rms sub_rms ");
	 	         	sql.append("  on sub_rms.rms_id = sub_ri.rms_id ");
	 	         	break;
	 	         case CATALOG_ITEM_REPORT:
	 	         	sql.append("inner join new_catalog_item nci ");
	 	         	sql.append("  on nci.item_id = item_iea.item_id ");
	 	         	sql.append("  and nci.warehouse_id = warehouse.warehouse_id ");
	 	         	break;
	 	       }
             
             sql.append("left outer join flc item_flc ");
             sql.append("  on item_flc.flc_id = item_ei.flc_id ");
             sql.append("left outer join mdc item_mdc ");
             sql.append("  on item_mdc.mdc_id = item_flc.mdc_id ");
             sql.append("left outer join ship_unit item_su ");
             sql.append("  on item_su.unit_id = item_iea.ship_unit_id ");
             sql.append("left outer join ejd_item_warehouse ptld ");
             sql.append("  on ptld.ejd_item_id = item_ei.ejd_item_id ");
             sql.append("  and ptld.warehouse_id = 1 ");
             sql.append("left outer join ejd_item_warehouse pitt ");
             sql.append("  on pitt.ejd_item_id = item_ei.ejd_item_id ");
             sql.append("  and pitt.warehouse_id = 2 ");
             
             //Where 
             sql.append("where item_type.itemtype = 'STOCK' ");
             switch (m_RptType) {
             	case  NEW_ITEM_REPORT:
             		sql.append("and ((trunc(item_ei.setup_date) >= to_date(?, 'mm/dd/yyyy') ");
             		sql.append("      and trunc(item_ei.setup_date) <= to_date(?, 'mm/dd/yyyy') ");
             		sql.append("      and item_eiw.disp_id = '1') ");
             		sql.append("    or (trunc(item_eiw.active_begin) >= to_date(?, 'mm/dd/yyyy') ");
             		sql.append("      and trunc(item_eiw.active_begin) <= to_date(?, 'mm/dd/yyyy') ");
             		sql.append("      and item_eiw.disp_id = '1') ");
             		sql.append("    or item_ei.ejd_item_id = ");
             		sql.append("        (select distinct eidh.ejd_item_id ");
             		sql.append("         from ejd_item_disp_history eidh ");
             		sql.append("         where eidh.ejd_item_id = item_eiw.ejd_item_id ");
             		sql.append("         and eidh.warehouse_id = item_eiw.warehouse_id ");
             		sql.append("         and eidh.new_disp_id = 1 and eidh.old_disp_id <> 1 ");
             		sql.append("         and (trunc(eidh.change_date) >= to_date(?, 'mm/dd/yyyy') ");
             		sql.append("              and (trunc(eidh.change_date) <= to_date(?, 'mm/dd/yyyy')) ");
             		sql.append("              and eidh.change_date = ");
             		sql.append("                (select max(lastchg.change_date) ");
             		sql.append("                 from ejd_item_disp_history lastchg ");
             		sql.append("                 where lastchg.ejd_item_id = item_eiw.ejd_item_id ");
             		sql.append("                 and lastchg.warehouse_id = item_eiw.warehouse_id) ");
             		sql.append("             ) ");
             		sql.append("        ) ");
             		sql.append("     ) ");
             		break;
	 	         case DISCONTINUED_ITEM_REPORT:
	 	         	sql.append("and item_ei.ejd_item_id = ");
	 	         	sql.append("    (select distinct eidh.ejd_item_id ");
	 	         	sql.append("     from ejd_item_disp_history eidh ");
	 	         	sql.append("     where eidh.ejd_item_id = item_eiw.ejd_item_id ");
	 	         	sql.append("     and eidh.warehouse_id = item_eiw.warehouse_id ");
	 	         	sql.append("     and (eidh.new_disp_id = 2 or eidh.new_disp_id = 3) ");
	 	         	sql.append("     and trunc(eidh.change_date) >= to_date(?, 'mm/dd/yyyy') ");
	 	         	sql.append("     and trunc(eidh.change_date) <= to_date(?, 'mm/dd/yyyy') ");
	 	         	sql.append("     and eidh.change_date = ");
	 	         	sql.append("         (select max(lastchg.change_date) ");
	 	         	sql.append("          from ejd_item_disp_history lastchg ");
	 	         	sql.append("          where lastchg.ejd_item_id = item_eiw.ejd_item_id ");
	 	         	sql.append("          and lastchg.warehouse_id = item_eiw.warehouse_id ");
	 	         	sql.append("         ) ");
	 	         	sql.append("    ) ");
	 	         	break;
	 	         case CATALOG_ITEM_REPORT:
	 	            sql.append("and trunc(nci.new_cat_date) >= to_date(?, 'mm/dd/yyyy') ");
	 	            sql.append("and trunc(nci.new_cat_date) <= to_date(?, 'mm/dd/yyyy') ");
	 	            break;
	 	       }
             
             //Order by
             switch (m_RptType) {
	 	         case  NEW_ITEM_REPORT:
	 	         	sql.append("order by item_ed.dept_num, item_vendor.name, item_iea.item_id, warehouse.warehouse_id ");
	 	         	break;
	 	         case DISCONTINUED_ITEM_REPORT:
	 	         	sql.append("order by item_ed.dept_num, item_iea.vendor_id, item_iea.item_id, warehouse.warehouse_id ");
	 	         	break;
	 	         case CATALOG_ITEM_REPORT:
	 	         	sql.append("order by item_iea.item_id, warehouse.warehouse_id ");
	 	         	break;
             }
            
            m_RptData = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[NewDiscItem]", ex);
         }
         
         finally {
         	sql = null;
         }         
      }
      else
         log.error("[NewDiscItem] prepareStatements - null db connection");
      
      return isPrepared;
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private int createCatalogItemCaptions()
   {
      int rowNum = 0;
      int col = 0;
      int m_CharWidth = 295;
      String title = String.format("New Catalog Items Report for %s - %s", m_BegDate, m_EndDate);
      
      //
      // creates Excel title
      addRow(rowNum++);
      addCell(col, title, m_StyleHdrLeft);
      rowNum++;
      
      //
      // Add the captions
      addRow(rowNum++);
      m_Sheet.setColumnWidth(col, (8 * m_CharWidth));
      addCell(col, "Item Number", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Description", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "Buyer Dept", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (9 * m_CharWidth));
      addCell(col, "Vendor#", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Vendor", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "NRHA", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "FLC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "UPC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "NBC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Stock Pack", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Unit", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "(Ptld) Catalog", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "(Pitt) Catalog", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Emery Cost", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Base Cost", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "RetailC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "SOQ Comments", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Setup Date", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Status Date", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "QTY", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Disposition", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Warehouse", m_StyleCaption);
      
      return rowNum;
      
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private int createDiscontinuedItemCaptions()
   {
      int rowNum = 0;
      int col = 0;
      int m_CharWidth = 295;
      String title = String.format("Discontinued Items Report for %s - %s", m_BegDate, m_EndDate);
      
      //
      // creates Excel title
      addRow(rowNum++);
      addCell(col, title, m_StyleHdrLeft);
      rowNum++;
      
      //
      // Add the captions
      addRow(rowNum++);
      m_Sheet.setColumnWidth(col, (4 * m_CharWidth));
      addCell(col, "Byr", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (9 * m_CharWidth));
      addCell(col, "Vendor#", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Vendor Name", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Item #", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Item Description", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "NRHA", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "FLC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "UPC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "NBC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "Stk Pk", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Pkg", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "RMS #", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (30 * m_CharWidth));
      addCell(col, "RMS Description", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Emery Cost", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Base", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "C Mkt", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (70 * m_CharWidth));
      addCell(col, "Comments", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Catalog Page", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "Auto Sub", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Suggested Sub", m_StyleCaption);

      m_Sheet.setColumnWidth(++col, (4 * m_CharWidth));
      addCell(col, "Byr", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "NRHA", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Vendor#", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Vendor Name", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "FLC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Item Description", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "UPC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "NBC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Stk Pk", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Emery Cost", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Base", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "C Mkt", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Pkg", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "RMS #", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (30 * m_CharWidth));
      addCell(col, "RMS Description", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "Disposition", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Warehouse", m_StyleCaption);
      
      return rowNum;
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private int createNewItemCaptions()
   {
   	int rowNum = 0;
      int col = 0;
      int m_CharWidth = 295;
      String title = String.format("New Items Report for %s - %s", m_BegDate, m_EndDate);
      
      //
      // creates Excel title
      addRow(rowNum++);
      addCell(col, title, m_StyleHdrLeft);
      rowNum++;
      
      //
      // Add the captions
      addRow(rowNum++);
      m_Sheet.setColumnWidth(col, (4 * m_CharWidth));
      addCell(col, "Byr", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Vendor Name", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Item #", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "NRHA", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "FLC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Item Description", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Base", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "C Mkt", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "UPC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Pkg", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "Stk Pk", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "NBC", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Comments", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "(Ptld)", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "(Pitt)", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Setup Date", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Status Date", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Active Begin", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Quantity", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "Disposition", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Vendor#", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Emery Cost", m_StyleCaption);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Warehouse", m_StyleCaption);
      
      return rowNum;
   }
   
   
   /*For debugging locally only.*/
/*
   public static void main(String args[]) throws FileNotFoundException, SQLException 
   {
      org.apache.log4j.BasicConfigurator.configure();
      NewDiscItem ndi = new NewDiscItem();
      Param[] parms = new Param[] {
            new Param("date", "04/01/2018", "startdate"), 
            new Param("date", "04/30/2018", "enddate"),
            new Param("integer", "0", "warehouse"),
            new Param("integer", "0", "rpttype")
      };
      
      ArrayList<Param> parmslist = new ArrayList<Param>();
      for (Param p : parms) {
         parmslist.add(p);
      }
      
      ndi.setParams(parmslist);
      java.util.Properties connProps = new java.util.Properties();
      connProps.put("user", "ejd");
      connProps.put("password", "boxer");

      ndi.m_Status = RptServer.RUNNING;
      ndi.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen", connProps);
      ndi.m_FilePath = "C:/Users/BCornwell/temp/";
      ndi.createReport();
    
      System.out.println("done");
   }
 	*/ 	
}
