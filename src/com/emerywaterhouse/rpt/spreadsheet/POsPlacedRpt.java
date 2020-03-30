/**
 * File: POReport.java
 * Description: PO report
 *
 * @author Eric Brownewell
 * <p>
 * Create Date: 08/19/2014
 * Last Update: $Id: POsPlacedRpt.java,v 1.6 2015/02/26 15:44:19 jfisher Exp $
 * <p>
 * History
 * <p>
 * $Log: POsPlacedRpt.java,v $
 * Revision 1.6  2015/02/26 15:44:19  jfisher
 * Fixed the bug that resulted from fixing the bug the resulted from fixing the bug.  Or - because I'm stupid.
 * <p>
 * Revision 1.5  2015/02/26 15:35:40  jfisher
 * Fixed the bug that resulted from fixing the bug.
 * <p>
 * Revision 1.4  2015/02/26 15:24:09  jfisher
 * Changed query to use materialized views and fixed a bug.
 * <p>
 * Revision 1.3  2014/10/27 19:21:14  ebrownewell
 * added forecast for last 90 days
 * <p>
 * Revision 1.2  2014/10/20 16:26:50  ebrownewell
 * New report for Dean Frost.
 * <p>
 * Revision 1.1  2014/08/22 14:59:35  ebrownewell
 * POs Placed report generation
 * updated sql query
 * adjusted buildOutputFile() to use proper data
 * adjusted createCaptions() to include proper names
 * adjusted setParams() to include the proper file name
 */
package com.emerywaterhouse.rpt.spreadsheet;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class POsPlacedRpt extends Report {
    private static final short MAX_COLS = 14;

    private PreparedStatement m_poData;

    //
    // The cell styles for each of the base columns in the spreadsheet.
    private XSSFCellStyle[] m_CellStyles;

    //
    // workbook entries.
    private XSSFWorkbook m_Wrkbk;
    private XSSFSheet m_Sheet;

    //parameters
    private String m_BegDate;
    private String m_EndDate;

    public POsPlacedRpt() {
        super();
        m_Wrkbk = new XSSFWorkbook();
        m_Sheet = m_Wrkbk.createSheet();
        setupWorkbook();
    }

    /**
     * Cleanup any allocated resources.
     */
    @Override
    public void finalize() throws Throwable {
        if (m_CellStyles != null) {
            for (int i = 0; i < m_CellStyles.length; i++)
                m_CellStyles[i] = null;
        }

        m_Sheet = null;
        m_Wrkbk = null;
        m_CellStyles = null;
        super.finalize();
    }


    /**
     * Executes the queries and builds the output file
     *
     * @throws java.io.FileNotFoundException
     */
    private boolean buildOutputFile() throws FileNotFoundException {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        XSSFRow row = null;
        short rowNum = 0;
        int colNum = 0;
        FileOutputStream outFile = null;
        ResultSet poData = null;
        boolean result = false;

        outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
        try {
            rowNum = createCaptions();

            m_poData.setDate(1, new java.sql.Date(dateFormat.parse(m_BegDate).getTime()));
            m_poData.setDate(2, new java.sql.Date(dateFormat.parse(m_EndDate).getTime()));
            poData = m_poData.executeQuery();

            while (poData.next() && m_Status == RptServer.RUNNING) {
                setCurAction("Building output file");

                row = createRow(rowNum++, MAX_COLS);
                colNum = 0;

                //TODO: potentially rearrange/rename when dean sends an updated excel doc.
                if (row != null) {
                    row.getCell(colNum++).setCellValue(poData.getString("whs")); //col 0
                    row.getCell(colNum++).setCellValue(poData.getString("vendor_name")); //col 1
                    row.getCell(colNum++).setCellValue(poData.getString("po_nbr")); //col 2
                    row.getCell(colNum++).setCellValue(poData.getString("item_nbr")); //col 3
                    row.getCell(colNum++).setCellValue(poData.getString("descr")); //col 4
                    row.getCell(colNum++).setCellValue(poData.getString("qty_ordered")); //col 5
                    row.getCell(colNum++).setCellValue(poData.getDouble("dollars_purchased")); //col 6
                    row.getCell(colNum++).setCellValue(poData.getString("on_hand")); //col 7
                    row.getCell(colNum++).setCellValue(poData.getString("ninety_day_sales")); //col 8
                    row.getCell(colNum++).setCellValue(poData.getString("velocity_cd")); //col 9
                    row.getCell(colNum++).setCellValue(poData.getString("fcst_qty")); //col 10
                    row.getCell(colNum++).setCellValue(poData.getString("last_90_days_fcst")); //col 11
                    String str = poData.getString("due_in_date");
                    if (str.indexOf(' ') != -1) {
                        str = str.substring(0, str.indexOf(' '));
                    }
                    row.getCell(colNum++).setCellValue(str); //col 12
                    row.getCell(colNum++).setCellValue(poData.getString("buying_dept")); //col 13
                }
            }

            m_Wrkbk.write(outFile);
            poData.close();

            result = true;
        } catch (Exception ex) {
            m_ErrMsg.append("Your report had the following errors: \r\n");
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());

            log.fatal("[POsPlacedRpt]:", ex);
        } finally {
            row = null;

            try {
                outFile.close();
            } catch (Exception e) {
                log.error("[POsPlacedRpt]:" + e);
            }

            outFile = null;
        }

        return result;
    }

    /**
     * Closes all the sql statements so they release the db cursors.
     */
    private void closeStatements() {
        closeStmt(m_poData);
    }

    /**
     * Sets the captions on the report.
     */
    private short createCaptions() {
        XSSFCellStyle captionStyle;
        captionStyle = m_Wrkbk.createCellStyle();
        captionStyle.setAlignment(HorizontalAlignment.CENTER);
        captionStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        captionStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFRow row = null;
        short rowNum = 0;
        int colNum = 0;

        if (m_Sheet == null)
            return 0;

        //
        // Create the row for the captions.
        row = m_Sheet.createRow(rowNum);

        if (row != null) {
            for (int i = 0; i < MAX_COLS; i++) {
                row.createCell(i);
            }
        }

        //TODO: potentially rearrange later when Dean sends revised excel doc.

        row.getCell(colNum).setCellValue("whs"); //col 0
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 4000);
        row.getCell(colNum).setCellValue("vendor_name"); //col 1
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 10000);
        row.getCell(colNum).setCellValue("po_nbr"); //col 2
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 4000);
        row.getCell(colNum).setCellValue("item_nbr"); //col 3
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 4000);
        row.getCell(colNum).setCellValue("descr"); //col 4
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 20000);
        row.getCell(colNum).setCellValue("qty_ordered"); //col 5
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("dollars_purchased"); //col 6
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 4000);
        row.getCell(colNum).setCellValue("on_hand"); //col 7
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("ninety_day_sales"); //col 8
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("velocity_cd"); //col 9
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("fcst_qty"); //col 10
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("fcst_qty_last_90"); //col 11
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("due_in_date"); //col 12
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 3000);
        row.getCell(colNum).setCellValue("buying_group"); //col 13
        row.getCell(colNum).setCellStyle(captionStyle);
        m_Sheet.setColumnWidth(colNum++, (short) 4000);

        return ++rowNum;
    }

    /**
     * @see com.emerywaterhouse.rpt.server.Report#createReport()
     */
    @Override
    public boolean createReport() {
        boolean created = false;
        m_Status = RptServer.RUNNING;

        try {
            m_EdbConn = m_RptProc.getEdbConn();

            if (prepareStatements())
                created = buildOutputFile();
        } catch (Exception ex) {
            log.fatal("[POsPlacedRpt]:", ex);
        } finally {
            closeStatements();

            if (m_Status == RptServer.RUNNING)
                m_Status = RptServer.STOPPED;
        }

        return created;
    }

    /**
     * Creates a row in the worksheet.
     * @param rowNum The row number.
     * @param colCnt The number of columns in the row.
     *
     * @return The formatted row of the spreadsheet.
     */
    private XSSFRow createRow(short rowNum, short colCnt) 
    {
        XSSFRow row = null;
        XSSFCell cell = null;

        if (m_Sheet == null)
            return row;

        row = m_Sheet.createRow(rowNum);

        //
        // set the type and style of the cell.
        if (row != null) {
            for (int i = 0; i < colCnt; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(m_CellStyles[i]);
            }
        }

        return row;
    }

    /**
     * Prepares the sql queries for execution.
     */
    private boolean prepareStatements() 
    {
        StringBuffer sql = new StringBuffer();
        boolean isPrepared = false;

        if (m_EdbConn != null) {
            try {
                sql.append("select ");
                sql.append("   whs, po_hdr.vendor_name, po_hdr.po_nbr, whs_po.item_nbr, whs_po.descr, whs_po.qty_ordered, ");
                sql.append("   whs_po.dollars_purchased, whs_po.on_hand, whs_po.ninety_day_sales, whs_po.velocity_cd, ");
                sql.append("   po_hdr.due_in_date, whs_po.buying_dept, fcst_12_week_mv.fcst_qty, ");
                sql.append("  fcst_prev_12_week_mv.fcst_qty last_90_days_fcst ");
                sql.append("from ");
                sql.append("   po_hdr ");
                sql.append("join ( ");
                sql.append("   select ");
                sql.append("      'PITTSTON' whs, po_dtl.po_nbr, po_dtl.item_nbr, po_dtl.descr, po_dtl.qty_ordered, ");
                sql.append("      (po_dtl.qty_ordered * po_dtl.emery_cost) dollars_purchased, iw.avail_qty on_hand, ");
                sql.append("      ninety_day_sales, po_dtl.velocity_cd, buying_dept ");
                sql.append("   from po_dtl ");
                sql.append("   join item_warehouse iw on iw.item_id = po_dtl.item_nbr and iw.warehouse_id = 2 ");
                sql.append("      ");
                sql.append("   join ( ");
                sql.append("      select sale_item.buying_dept, sale_item.item_id, sum(qty_shipped) ninety_day_sales ");
                sql.append("      from sale_item ");
                sql.append("      where ");
                sql.append("         sale_item.invoice_date >= trunc(sysdate) - 90 and tran_type = 'SALE' and ");
                sql.append("         sale_item.warehouse = 'PITTSTON' and sale_item.item_id is not null ");
                sql.append("      group by sale_item.item_id, sale_item.buying_dept ");
                sql.append("   ");
                sql.append("   ) sales on sales.item_id = po_dtl.item_nbr ");
                sql.append("   where po_dtl.warehouse = '04' ");
                sql.append("   union ");
                sql.append("   select ");
                sql.append("      'PORTLAND' whs, po_dtl.po_nbr, po_dtl.item_nbr, po_dtl.descr, po_dtl.qty_ordered, ");
                sql.append("      (po_dtl.qty_ordered * po_dtl.emery_cost) dollars_purchased, iw.avail_qty on_hand, ");
                sql.append("      ninety_day_sales, po_dtl.velocity_cd, buying_dept ");
                sql.append("   from po_dtl ");
                sql.append("   join item_warehouse iw on iw.item_id = po_dtl.item_nbr and iw.warehouse_id = 1 ");
                sql.append("   join ( ");
                sql.append("      select sale_item.buying_dept, sale_item.item_id, sum(qty_shipped) ninety_day_sales ");
                sql.append("      from sale_item ");
                sql.append("      where ");
                sql.append("         sale_item.invoice_date >= trunc(sysdate) - 90 and tran_type = 'SALE' and ");
                sql.append("         sale_item.warehouse = 'PORTLAND' and sale_item.item_id is not null ");
                sql.append("      group by sale_item.item_id, sale_item.buying_dept ");
                sql.append("   ");
                sql.append("   ) sales on sales.item_id = po_dtl.item_nbr ");
                sql.append("   where po_dtl.warehouse = '01' ");
                sql.append(") whs_po on whs_po.po_nbr = po_hdr.po_nbr ");
                sql.append("left outer join fcst_12_week_mv on fcst_12_week_mv.whs_name = whs_po.whs and ");
                sql.append("   fcst_12_week_mv.prod_no = whs_po.item_nbr ");
                sql.append("left outer join fcst_prev_12_week_mv on fcst_prev_12_week_mv.whs_name = whs_po.whs and ");
                sql.append("   fcst_prev_12_week_mv.prod_no = whs_po.item_nbr ");
                sql.append("where po_hdr.po_date between ? and ? ");
                sql.append("order by po_hdr.po_nbr ");

                m_poData = m_EdbConn.prepareStatement(sql.toString());

                isPrepared = true;
            } 
            catch (SQLException ex) {
                log.error("[POsPlacedRpt]:", ex);
            } 
            
            finally {
                sql = null;
            }
        } 
        else
            log.error("[POsPlacedRpt].prepareStatements - null oracle connection");

        return isPrepared;
    }

    /**
     * Sets the parameters of this report.
     * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
     */
    @Override
    public void setParams(ArrayList<Param> params) 
    {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        StringBuffer fileName = new StringBuffer();
        String tmp = Long.toString(System.currentTimeMillis());
        Param param = null;

        fileName.append("Purchase_Orders_Placed");
        fileName.append("-");
        fileName.append(tmp.substring(tmp.length() - 5, tmp.length()));
        fileName.append(".xlsx");
        m_FileNames.add(fileName.toString());

        for (int i = 0; i < params.size(); i++) {
            param = params.get(i);

            if (param.name.equalsIgnoreCase("begdate"))
                m_BegDate = param.value;

            if (param.name.equalsIgnoreCase("enddate"))
                m_EndDate = param.value;

        }

        if ((m_BegDate == null || m_BegDate.length() == 0) || (m_EndDate == null || m_EndDate.length() == 0)) {
            String[] defaultDates = setDefaultDates();
            m_BegDate = defaultDates[0];
            m_EndDate = defaultDates[1];
        }

        try {
            Date begDate = new SimpleDateFormat("M/dd/yyyy").parse(m_BegDate);
            Date endDate = new SimpleDateFormat("M/dd/yyyy").parse(m_EndDate);
            m_BegDate = dateFormat.format(begDate);
            m_EndDate = dateFormat.format(endDate);
        } 
        
        catch (ParseException e) {
            try {
                Date begDate = dateFormat.parse(m_BegDate);
                Date endDate = dateFormat.parse(m_EndDate);
                m_BegDate = dateFormat.format(begDate);
                m_EndDate = dateFormat.format(endDate);
            } catch (ParseException ex) {
                log.error("[POsPlacedRpt]ParseException:", ex);
            }
        }

    }

    private static String[] setDefaultDates() 
    {
        //set default begin and end dates in case we dont get them.
        String[] results = new String[2];
        Date now = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        Date weekAgo = getDateInPast(now, 7);
        results[0] = dateFormat.format(weekAgo);
        results[1] = dateFormat.format(now);

        return results;
    }

    /**
     * Internal utility function to get a date from a supplied number of days ago.
     * @param date The date to use as a reference.
     * @param numDays The number of days to set the date back.
     *
     * @return a date object that is numDays before the date that was passed in
     */
    private static Date getDateInPast(final Date date, int numDays) 
    {
        Date result = new Date(date.getTime());
        GregorianCalendar calendar = new GregorianCalendar();
        calendar.setTime(result);
        calendar.add(Calendar.DATE, -numDays);
        result.setTime(calendar.getTime().getTime());

        return result;
    }

    /**
     * Sets up the styles for the cells based on the column data.  Does any other inititialization
     * needed by the workbook.
     */
    private void setupWorkbook() 
    {
        XSSFCellStyle styleText;      // Text left justified
        XSSFCellStyle styleInt;       // Style with 0 decimals
        XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)

        styleText = m_Wrkbk.createCellStyle();
        styleText.setAlignment(HorizontalAlignment.LEFT);

        styleInt = m_Wrkbk.createCellStyle();
        styleInt.setAlignment(HorizontalAlignment.RIGHT);
        styleInt.setDataFormat((short) 3);

        styleMoney = m_Wrkbk.createCellStyle();
        styleMoney.setAlignment(HorizontalAlignment.RIGHT);
        styleMoney.setDataFormat((short) 8);

        m_CellStyles = new XSSFCellStyle[]{
                styleText,   // col 0 whs
                styleText,   // col 1 vendor_name
                styleText,   // col 2 po_nbr
                styleText,   // col 3 item_nbr
                styleText,   // col 4 descr
                styleText,   // col 5 qty_ordered
                styleMoney,  // col 6 dollars_purchased
                styleText,   // col 7 on_hand
                styleText,   // col 8 ninety_day_sales
                styleText,   // col 9 velocity_cd
                styleText,   // col 10 fcst_qty
                styleText,   // col 11 fcst_qty_last_90
                styleText,   // col 12 due_in_date
                styleText,   // col 13 Buying Category
        };

        styleText = null;
        styleInt = null;
        styleMoney = null;
    }
}
