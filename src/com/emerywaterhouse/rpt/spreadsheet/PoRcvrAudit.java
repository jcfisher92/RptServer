package com.emerywaterhouse.rpt.spreadsheet;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;
import org.apache.log4j.BasicConfigurator;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;

public class PoRcvrAudit extends Report {
    private static final short maxCols = 14;

    private PreparedStatement m_RcvrAuditRcvd;
    private PreparedStatement m_RcvrAuditOpen;

    //
    // The cell styles for each of the columns in the spreadsheet.
    private XSSFCellStyle[] m_CellStyles;
    private XSSFCellStyle[] m_CellStylesEx;

    private CellStyle rightAlign;
    private CellStyle leftAlign;
    private CellStyle centerAlign;

    //
    // workbook entries.
    private XSSFWorkbook m_Wrkbk;
    private XSSFSheet m_Sheet;

    private int emRcvrNum;

    private boolean printedTotals = true; //start out true so we don't try to print totals on the first line

    DecimalFormat df = new DecimalFormat("#.00");

    /**
     * Default constructor
     */
    public PoRcvrAudit() {
        super();

        m_Wrkbk = new XSSFWorkbook();
        m_Sheet = m_Wrkbk.createSheet("PO Receiver Audit Report");

        createHeader(); //create the title, date, and page text.

        m_MaxRunTime = RptServer.HOUR * 12;

        setupWorkbook();
    }

    /**
     * Cleanup any allocated resources.
     */
    public void finalize() throws Throwable {
        if (m_CellStyles != null) {
            for (int i = 0; i < m_CellStyles.length; i++)
                m_CellStyles[i] = null;
        }

        if (m_CellStylesEx != null) {
            for (int i = 0; i < m_CellStylesEx.length; i++)
                m_CellStylesEx[i] = null;
        }

        m_Sheet = null;
        m_Wrkbk = null;
        m_CellStyles = null;
        m_CellStylesEx = null;
        m_RcvrAuditRcvd = null;
        m_RcvrAuditOpen = null;

        super.finalize();
    }

    /**
     * Executes the queries and builds the output file
     *
     * @throws FileNotFoundException if the output file cant be created.
     */
    private boolean buildOutputFile() throws Exception {
        LinkedList<Integer> emeryRcvrNumbers;


        FileOutputStream outFile;
        boolean result = false;
        int rowNum = 1;

        if (emRcvrNum == -1) {
            emeryRcvrNumbers = getAllEmeryReceiverNumbers();
        } else {
            emeryRcvrNumbers = new LinkedList<>();
            emeryRcvrNumbers.add(emRcvrNum);
        }

        if (emeryRcvrNumbers.isEmpty()) {
            log.error("Emery Receiver Numbers list is empty. Cannot create report.");
            throw new Exception("Emery Receiver Numbers list is empty. Cannot create report.");
        }

        String m_FilePath = "reports/";
        outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

        try {

            for (int rcvrNum : emeryRcvrNumbers) {
                XSSFRow row;
                int colNum;
                String msg = "processing receiver po record %s";
                ResultSet poData;
                String rcvrHdrId;
                String prevPoNbr = "";

                m_RcvrAuditRcvd.setInt(1, rcvrNum);

                poData = m_RcvrAuditRcvd.executeQuery();

                int totalQtyOpen = 0;
                int totalQtyExp = 0;
                int totalQtyReceived = 0;
                int totalQtyPut = 0;
                double totalExtFrt = 0.00;
                double totalextCost = 0.00;
                double totalExtTotCost = 0.00;
                int totalQtyDiff = 0;
                double totalCostDiff = 0.00;

                while (poData.next() && m_Status == RptServer.RUNNING) {
                    rcvrHdrId = poData.getString("rcvr_po_hdr_id");
                    setCurAction(String.format(msg, rcvrHdrId));

                    String remitId = poData.getString("remit_to_id");
                    String eisVndNum = poData.getString("eis_vnd_nbr");
                    String vndName = poData.getString("vnd_name");
                    String poNum = poData.getString("po_nbr");
                    String carrier = poData.getString("rcvr_carrier");
                    String invoiceTerms = poData.getString("po_inv_terms");
                    Date receiptDate = poData.getDate("receipt_date");
                    String emeryRcvrNum = poData.getString("emery_rcvr_nbr");
                    String fasRcvrNum = poData.getString("fascor_rcvr_nbr");

                    String buyer = poData.getString("buyer_name");
                    Date dateClosed = poData.getDate("date_closed");
                    String comments = poData.getString("rcvr_comments");

                    if (!prevPoNbr.equalsIgnoreCase(poNum)) { //if we're onto a different po number we want to create a new receiver header.
                        //first we want to print a row of totals, so take the total counts and print them in their correct columns.
                        if (!printedTotals) {
                            printTotals(rowNum, totalQtyOpen, totalQtyExp, totalQtyPut, totalQtyReceived,
                                    totalExtFrt, totalextCost, totalExtTotCost, totalQtyDiff, totalCostDiff);
                            rowNum++;

                            rowNum = writeOpenLines(rcvrNum, rowNum);

                            addPageBreak(rowNum++);
                        }

                        totalQtyReceived = 0;
                        totalExtFrt = 0.00;
                        totalextCost = 0.00;
                        totalExtTotCost = 0.00;
                        totalQtyDiff = 0;
                        totalCostDiff = 0.00;

                        rowNum = createReceiverHeader(rowNum++, remitId, eisVndNum, vndName, poNum,
                                carrier, invoiceTerms, receiptDate, emeryRcvrNum, fasRcvrNum, buyer, dateClosed, comments);

                        prevPoNbr = poNum; //set prevPoNbr to the current po number so we don't keep using the old one (and remaking headers)
                    }

                    row = createRow(rowNum++, maxCols, false);
                    colNum = 0;

                    if (row != null) {
                        int qtyOpen = poData.getInt("qty_ordered");
                        int qtyExp = poData.getInt("qty_expected");
                        int qtyReceived = poData.getInt("qty_received");
                        int qtyPut = poData.getInt("qty_put_away");
                        double extFrt = poData.getDouble("ext_freight");
                        double extCost = poData.getDouble("ext_cost");
                        double extTotCost = poData.getDouble("ext_tot_cost");
                        int qtyDiff = poData.getInt("diff_units");
                        double totCostDiff = poData.getDouble("diff_cost");

                        double unitCost = poData.getDouble("unit_cost");
                        double unitFreight = poData.getDouble("unit_freight");

                        totalQtyOpen += qtyOpen;
                        totalQtyExp += qtyExp;
                        totalQtyPut += qtyPut;
                        totalQtyReceived += qtyReceived;
                        totalExtFrt += extFrt;
                        totalextCost += extCost;
                        totalExtTotCost += extTotCost;
                        totalQtyDiff += qtyDiff;
                        totalCostDiff += totCostDiff;

                        row.getCell(colNum++).setCellValue(poData.getString("item_nbr")); //item id

                        Cell cell = row.getCell(colNum++);//we need to set the style for item desc to null so it doesnt get centered.
                        cell.setCellValue(poData.getString("item_desc")); //item description
                        cell.setCellStyle(leftAlign);

                        row.getCell(colNum++).setCellValue(poData.getString("shp_unit")); //ship unit
                        row.getCell(colNum++).setCellValue(String.valueOf(qtyOpen)); //qty Open
                        row.getCell(colNum++).setCellValue(String.valueOf(qtyExp)); //qty expected
                        row.getCell(colNum++).setCellValue(String.valueOf(qtyReceived)); //qty received
                        row.getCell(colNum++).setCellValue(String.valueOf(qtyPut)); //qty put away
                        row.getCell(colNum++).setCellValue(String.valueOf(df.format(unitCost))); //cost / ship unit
                        row.getCell(colNum++).setCellValue(String.valueOf(df.format(unitFreight))); //frt / ship unit
                        row.getCell(colNum++).setCellValue(String.valueOf(df.format(extCost))); //ext cost
                        row.getCell(colNum++).setCellValue(String.valueOf(df.format(extFrt))); //ext frt
                        row.getCell(colNum++).setCellValue(String.valueOf(df.format(extTotCost))); //ext total cost
                        row.getCell(colNum++).setCellValue(String.valueOf(qtyDiff)); //variance units
                        row.getCell(colNum).setCellValue(String.valueOf(df.format(totCostDiff))); //variance cost

                    }
                }

                //print the last summary row
                if (!printedTotals) {
                    printTotals(rowNum++, totalQtyOpen, totalQtyExp, totalQtyPut, totalQtyReceived, totalExtFrt, totalextCost,
                            totalExtTotCost, totalQtyDiff, totalCostDiff);

                    rowNum = writeOpenLines(rcvrNum, rowNum);

                    addPageBreak(rowNum++);
                }

                poData.close();
            }

            result = true;
            m_Wrkbk.write(outFile);

        } catch (Exception ex) {
            m_ErrMsg.append("Your report had the following errors: \r\n");
            m_ErrMsg.append(ex.getClass().getName()).append("\r\n");
            m_ErrMsg.append(ex.getMessage());

            log.fatal("exception:", ex);
        } finally {
            closeStatements();

            try {
                outFile.close();
            } catch (Exception e) {
                log.error(e);
            }
        }

        return result;
    }

    private void printTotals(int rownum, int totalQtyOpen, int totalQtyExp, int totalQtyPut, int totalQtyReceived, double totalExtFrt,
                             double totalExtCost, double totalExtTotCost, int totalQtyDiff, double totalCostDiff) 
    {
        XSSFRow row = createRow(rownum, maxCols, false);

        XSSFCell cell;

        CellStyle topLine;
        topLine = m_Wrkbk.createCellStyle();
        topLine.setBorderTop(BorderStyle.MEDIUM);
        topLine.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(3); //qty open
        cell.setCellValue(String.valueOf(totalQtyOpen));
        cell.setCellStyle(topLine);

        cell = row.getCell(4); //qty exp
        cell.setCellValue(String.valueOf(totalQtyExp));
        cell.setCellStyle(topLine);

        cell = row.getCell(5); //qty rcvd
        cell.setCellValue(String.valueOf(totalQtyReceived));
        cell.setCellStyle(topLine);

        cell = row.getCell(6); //qty put
        cell.setCellValue(String.valueOf(totalQtyPut));
        cell.setCellStyle(topLine);

        cell = row.getCell(9); //extended cost
        cell.setCellValue(String.valueOf(df.format(totalExtCost)));
        cell.setCellStyle(topLine);

        cell = row.getCell(10); //extended freight
        cell.setCellValue(String.valueOf(df.format(totalExtFrt)));
        cell.setCellStyle(topLine);

        cell = row.getCell(11); //extended tot cost
        cell.setCellValue(String.valueOf(df.format(totalExtTotCost)));
        cell.setCellStyle(topLine);

        cell = row.getCell(12); //qty diff
        cell.setCellValue(String.valueOf(totalQtyDiff));
        cell.setCellStyle(topLine);

        cell = row.getCell(13); //total cost diff
        cell.setCellValue(String.valueOf(df.format(totalCostDiff)));
        cell.setCellStyle(topLine);

        printedTotals = true;

    }

    private void addPageBreak(int rownum){
        //m_Sheet.setRowBreak(rownum);
        XSSFRow row = createRow(rownum, 1, true);

        row.getCell(0).setCellValue("EMY-REPLACE WITH PAGE BREAK-EMY");

    }

    /**
     * Close all the open queries
     */
    private void closeStatements() {
        closeStmt(m_RcvrAuditOpen);
        closeStmt(m_RcvrAuditRcvd);

    }

    private int writeOpenLines(int rcvrNum, int rowNum) throws SQLException {

        rowNum++; //give us a space between the total row and the open rows.

        m_RcvrAuditOpen.setInt(1, rcvrNum);

        XSSFRow row;
        int colNum;

        try(ResultSet rs = m_RcvrAuditOpen.executeQuery()){
            while(rs.next()){
                row = createRow(rowNum++, maxCols, false);
                colNum = 0;

                if (row != null) {
                    int qtyOpen = rs.getInt("qty_ordered");
                    int qtyExp = rs.getInt("qty_expected");
                    int qtyReceived = rs.getInt("qty_received");
                    int qtyPut = rs.getInt("qty_put_away");
                    double extFrt = rs.getDouble("ext_freight");
                    double extCost = rs.getDouble("ext_cost");
                    double extTotCost = rs.getDouble("ext_tot_cost");
                    int qtyDiff = rs.getInt("diff_units");
                    double totCostDiff = rs.getDouble("diff_cost");

                    double unitCost = rs.getDouble("unit_cost");
                    double unitFreight = rs.getDouble("unit_freight");

                    row.getCell(colNum++).setCellValue(rs.getString("item_nbr")); //item id

                    Cell cell = row.getCell(colNum++);
                    cell.setCellValue(rs.getString("item_desc")); //item description
                    cell.setCellStyle(leftAlign);

                    row.getCell(colNum++).setCellValue(rs.getString("shp_unit")); //ship unit
                    row.getCell(colNum++).setCellValue(String.valueOf(qtyOpen)); //qty Open
                    row.getCell(colNum++).setCellValue(String.valueOf(qtyExp)); //qty expected
                    row.getCell(colNum++).setCellValue(String.valueOf(qtyReceived)); //qty received
                    row.getCell(colNum++).setCellValue(String.valueOf(qtyPut)); //qty put away
                    row.getCell(colNum++).setCellValue(String.valueOf(df.format(unitCost))); //cost / ship unit
                    row.getCell(colNum++).setCellValue(String.valueOf(df.format(unitFreight))); //frt / ship unit
                    row.getCell(colNum++).setCellValue(String.valueOf(df.format(extCost))); //ext cost
                    row.getCell(colNum++).setCellValue(String.valueOf(df.format(extFrt))); //ext frt
                    row.getCell(colNum++).setCellValue(String.valueOf(df.format(extTotCost))); //ext total cost
                    row.getCell(colNum++).setCellValue(String.valueOf(qtyDiff)); //variance units
                    row.getCell(colNum).setCellValue(String.valueOf(df.format(totCostDiff))); //variance cost
                }

            }
        }

        return rowNum;
    }

    /**
     * Sets the captions on the report.
     */
    private int createCaptions(int rowNum) {
        XSSFCellStyle csCaption;
        XSSFCell cell;
        XSSFRow row;
        int colNum = 0;
        short rowHeight = 1000;

        csCaption = m_Wrkbk.createCellStyle();
        csCaption.setAlignment(HorizontalAlignment.CENTER);
        csCaption.setWrapText(true);

        if (m_Sheet != null) {
            // Create the row for the captions.
            row = m_Sheet.createRow(rowNum++);
            row.setHeight(rowHeight);

            for (int i = 0; i < maxCols; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(csCaption);
            }

            row.getCell(colNum).setCellValue("Item #");
            m_Sheet.setColumnWidth(colNum++, 3039);

            row.getCell(colNum).setCellValue("Item Description");
            m_Sheet.setColumnWidth(colNum++, 4500);

            row.getCell(colNum).setCellValue("Ship Unit");
            m_Sheet.setColumnWidth(colNum++, 2500);

            row.getCell(colNum).setCellValue("Qty Open");
            m_Sheet.setColumnWidth(colNum++, 1500);

            row.getCell(colNum).setCellValue("Qty Exp");
            m_Sheet.setColumnWidth(colNum++, 2000);

            row.getCell(colNum).setCellValue("Qty Rcvd");
            m_Sheet.setColumnWidth(colNum++, 1500);

            row.getCell(colNum).setCellValue("Qty Put");
            m_Sheet.setColumnWidth(colNum++, 1500);

            row.getCell(colNum).setCellValue("Unit Cost");
            m_Sheet.setColumnWidth(colNum++, 2000);

            row.getCell(colNum).setCellValue("Unit Freight");
            m_Sheet.setColumnWidth(colNum++, 1700);

            row.getCell(colNum).setCellValue("Ext Cost");
            m_Sheet.setColumnWidth(colNum++, 2800);

            row.getCell(colNum).setCellValue("Ext Freight");
            m_Sheet.setColumnWidth(colNum++, 2000);

            row.getCell(colNum).setCellValue("Ext Total Cost");
            m_Sheet.setColumnWidth(colNum++, 2500);

            row.getCell(colNum).setCellValue("Qty Diff");
            m_Sheet.setColumnWidth(colNum++, 2200);

            row.getCell(colNum).setCellValue("Total Cost Diff");
            m_Sheet.setColumnWidth(colNum, 2000);
        }

        return rowNum;
    }

    private int createReceiverHeader(int rowNum, String remitId, String eisVndNum, String vndName, String poNum, String carrier, String invoiceTerms,
                                     Date receiptDate, String emeryRcvrNum, String fascorRcvrNum, String buyer, Date dateClosed, String comments) {
        XSSFRow row;

        row = createRow(++rowNum, maxCols, true);

        row.getCell(4).setCellValue("Emery Rcvr Nbr: " + emeryRcvrNum);


        rowNum++;
        row = createRow(++rowNum, maxCols, true);

        row.getCell(0).setCellValue("PO Nbr: ");
        row.getCell(1).setCellValue(poNum);

        row.getCell(2).setCellValue("Vendor ID");
        row.getCell(3).setCellValue(eisVndNum);

        row.getCell(5).setCellValue("Vendor: ");
        row.getCell(6).setCellValue(vndName);


        rowNum++;
        row = createRow(rowNum, maxCols, true);

        row.getCell(0).setCellValue("Fas Rcvr: ");
        row.getCell(1).setCellValue(fascorRcvrNum);

        row.getCell(2).setCellValue("Remit ID: ");
        row.getCell(3).setCellValue(remitId);

        row.getCell(5).setCellValue("Buyer: ");
        row.getCell(6).setCellValue(buyer);

        row.getCell(11).setCellValue("Invoice Terms: ");
        row.getCell(11).setCellStyle(rightAlign);
        row.getCell(12).setCellValue(invoiceTerms);

        rowNum++;
        row = createRow(rowNum, maxCols, true);

        row.getCell(0).setCellValue("Date Closed: ");
        if (dateClosed != null)
            row.getCell(1).setCellValue(dateClosed.toString());
        else
            row.getCell(1).setCellValue("");

        row.getCell(5).setCellValue("Receipt Date: ");
        row.getCell(5).setCellStyle(rightAlign);
        row.getCell(6).setCellValue(receiptDate.toString());

        row.getCell(11).setCellValue("Carrier: ");
        row.getCell(11).setCellStyle(rightAlign);
        row.getCell(12).setCellValue(carrier);

        rowNum++;
        row = createRow(rowNum, maxCols, true);

        row.getCell(1).setCellValue("Fascor Comment: ");
        row.getCell(1).setCellStyle(rightAlign);
        row.getCell(2).setCellValue(comments);

        rowNum++;

        createCaptions(rowNum++);

        printedTotals = false;

        return rowNum;
    }

    private void createHeader() {
        String date = new SimpleDateFormat("dd/MM/yyyy").format(new Date());

        Header header = m_Sheet.getHeader();
        header.setLeft("Receiver Audit Report");
        header.setRight(date + " Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

    }

    private LinkedList<Integer> getAllEmeryReceiverNumbers() throws SQLException {
        LinkedList<Integer> res = new LinkedList<>();

        String sql = "select rcvr_po_hdr.review_purch, rcvr_po_hdr.emery_rcvr_nbr, vendor.name\n" +
                "from rcvr_po_hdr\n" +
                "join vendor on vendor.vendor_id = rcvr_po_hdr.vendor_id\n" +
                "where\n" +
                "   vendor.vendor_id = rcvr_po_hdr.vendor_id and\n" +
                "   review_purch = 'N' and\n" +
                "   po_nbr not like 'TR%' and\n" +
                "   po_nbr not like 'RR%'\n" +
                "order by\n" +
                "   vendor.name\n";

        try (PreparedStatement stmt = m_EdbConn.prepareStatement(sql)) {
            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    res.add(rs.getInt("emery_rcvr_nbr"));
                }
            }
        }

        return res;
    }

    /**
     * @see Report#createReport()
     */
    @Override
    public boolean createReport() {
        boolean created = false;
        m_Status = RptServer.RUNNING;

        try {
            //m_EdbConn = m_RptProc.getEdbConn();

            //for testing purposes:
            //todo:remove before pushing to prod
            java.util.Properties conProps = new java.util.Properties();
            conProps.put("user", "ejd");
            conProps.put("password", "boxer");

            m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://10.128.0.11/emery_jensen", conProps);

            if (prepareStatements())
                created = buildOutputFile();
        } catch (Exception ex) {
            log.fatal("exception:", ex);
        } finally {
            if (m_Status == RptServer.RUNNING)
                m_Status = RptServer.STOPPED;
        }

        return created;
    }

    /**
     * Creates a row in the worksheet.
     *
     * @param rowNum The row number.
     * @param colCnt The number of columns in the row.
     * @return The formatted row of the spreadsheet.
     */
    private XSSFRow createRow(int rowNum, int colCnt, boolean ignoreStyle) {
        XSSFRow row = null;

        if (m_Sheet != null) {

            row = m_Sheet.createRow(rowNum);

            // set the type and style of the cell.
            for (int i = 0; i < colCnt; i++) {
                Cell cell = row.createCell(i);
                if (!ignoreStyle)
                    cell.setCellStyle(centerAlign);
            }

        }
        return row;
    }

    /**
     * Prepares the sql queries for execution.
     */
    private boolean prepareStatements() {
        StringBuilder sql = new StringBuilder();
        boolean isPrepared = false;

        if (m_EdbConn != null) {

            try {
                sql.setLength(0);
                sql.append("select * from po_receiver_view where emery_rcvr_nbr = ? and qty_expected is not null order by item_desc");

                m_RcvrAuditRcvd = m_EdbConn.prepareStatement(sql.toString());

                sql.setLength(0);
                sql.append("select * from po_receiver_view where emery_rcvr_nbr = ? and qty_expected is null order by item_desc");

                m_RcvrAuditOpen = m_EdbConn.prepareStatement(sql.toString());

                isPrepared = true;
            } catch (SQLException ex) {
                log.error("exception:", ex);
            }
        } else
            log.error("[RptServer#PoRcvrAudit]Null oracle or fascor connection");

        return isPrepared;
    }

    /**
     * Sets the parameters of this report.
     *
     * @see Report#setParams(ArrayList)
     */
    public void setParams(ArrayList<Param> params) {
        for (Param p : params) {
            switch (p.name) {
                case "EmRcvrNum":
                    if (p.value.equalsIgnoreCase("all")) {
                        emRcvrNum = -1;
                    } else {
                        emRcvrNum = Integer.parseInt(p.value);
                    }
                    break;
                case "Email":
                    break;
            }
        }

        SimpleDateFormat df = new SimpleDateFormat("MM-dd-yy");
        m_FileNames.add(String.format("poRcvrAudit-%s.xlsx", df.format(new Date())));
    }

    /**
     * Sets up the styles for the cells based on the column data.  Does any other initialization
     * needed by the workbook.
     * Note -
     * The styles came from the original Excel worksheet.
     */
    private void setupWorkbook() {
        m_Sheet.setMargin(XSSFSheet.BottomMargin, .75);
        m_Sheet.setMargin(XSSFSheet.TopMargin, .75);
        m_Sheet.setMargin(XSSFSheet.LeftMargin, 0.0);
        m_Sheet.setMargin(XSSFSheet.RightMargin, 0.0);
        m_Sheet.getPrintSetup().setLandscape(true);

        leftAlign = m_Wrkbk.createCellStyle();
        leftAlign.setAlignment(HorizontalAlignment.LEFT);

        rightAlign = m_Wrkbk.createCellStyle();
        rightAlign.setAlignment(HorizontalAlignment.RIGHT);

        centerAlign = m_Wrkbk.createCellStyle();
        centerAlign.setAlignment(HorizontalAlignment.CENTER);
        //centerAlign.setWrapText(true);
    }

    public static void main(String... args) {
        BasicConfigurator.configure();

        PoRcvrAudit rpt = new PoRcvrAudit();

        ArrayList<Param> params = new ArrayList<>();

        params.add(new Param("String", "All", "EmRcvrNum"));

        rpt.setParams(params);

        boolean created = rpt.createReport();

        System.out.println("Created? " + created);
    }

}
