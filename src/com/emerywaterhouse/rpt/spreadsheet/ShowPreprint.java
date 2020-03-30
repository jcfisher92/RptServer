/**
 * File: ShowPreprint.java
 * Description: Builds the Show Preprint data report (EIS calls this Show Order
 * Form which corresponds to ShowPreprint.pas).
 *
 * @author Naresh Pasnur
 * <p>
 * Create Date: 11/16/2011
 * Last Update: $Id: ShowPreprint.java,v 1.23 2015/02/26 15:50:55 ebrownewell Exp $
 * <p>
 * History:
 * $Log: ShowPreprint.java,v $
 * Revision 1.23  2015/02/26 15:50:55  ebrownewell
 * *** empty log message ***
 * <p>
 * Revision 1.22  2015/01/12 19:53:17  ebrownewell
 * Updated class to make room for cusomer stamps
 * <p>
 * Revision 1.21  2015/01/07 17:27:09  ebrownewell
 * updated margins
 * <p>
 * Revision 1.20  2014/10/20 16:26:07  ebrownewell
 * combined QB rows into one row, and redesigned the header to not take up so much room. Set header to repeat for every page.
 * <p>
 * Revision 1.19  2014/08/28 14:08:08  ebrownewell
 * fixed minor bugs
 * <p>
 * Revision 1.18  2014/08/22 20:40:30  ebrownewell
 * -added two StringBuffers to buildOutputFile named "priceBuffer" and "qtyBuffer", these now store the results for column 4 and column 5, append a new line after each result, and then set the cell value to the StringBuffer. This cuts out the rows that are empty except for the QB results, and adds them all into one cell but on multiple lines.
 * -commented out some old code that is no longer needed after the update
 * -added vendor name and a "page# of totapages#" to the right section of the footer.
 * <p>
 * Revision 1.17  2014/03/17 18:37:50  epearson
 * updated characters to UTF-8
 * <p>
 * Revision 1.16  2013/02/05 20:55:07  npasnur
 * Added code to include packet header information.
 * <p>
 * Revision 1.15  2013/02/04 17:10:10  jfisher
 * Removed some debugging logs.
 * <p>
 * Revision 1.14  2013/02/04 17:03:38  jfisher
 * Fixed a bug in the report where it wasn't correctly returning a true value in the buildOutputFile method.
 * <p>
 * Revision 1.13  2013/01/16 15:11:06  jfisher
 * Added CVS tag and brought logging up to specs.
 * <p>
 * Revision 1.12  2013/01/09 16:05:21  npasnur
 * Added forward slash to the file/image path
 * <p>
 * Revision 1.11  2012/12/28 17:49:16  npasnur
 * Removed hard-coded reference to G drive
 * <p>
 * Revision 1.10  2012/07/04 23:53:23  npasnur
 * Few changes and cleaned up unnecessary code
 */
package com.emerywaterhouse.rpt.spreadsheet;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;
import org.apache.commons.io.IOUtils;
import org.apache.log4j.BasicConfigurator;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.Enumeration;
import java.util.Vector;


public class ShowPreprint extends Report {
    private HSSFSheet m_Sheet;
    private String m_ShowDesc;                 // Description (name) of show
    private String m_PromoId;                 // Show promotion identifier
    private String m_PacketId;                // Show packet identifier
    private String m_PacketFooter;            // Show packet identifier
    private boolean m_ExtraEntry;
    private boolean m_Hotbuy;
    private boolean m_Incredible;
    private boolean m_NoBrkCarton;
    private boolean m_SlamDunk;
    protected String m_ImagePath;              // The path to the emery logo
    private String m_CustomerId;

    private PreparedStatement m_StmtShowData;
    private PreparedStatement m_StmtShowVendors;
    private PreparedStatement m_StmtShowPromoHdr;
    private PreparedStatement m_StmtQtBuys;
    private PreparedStatement m_StmtShowPacketHdr;

    private HSSFWorkbook m_Wrkbk;

    private static short BASE_COLS = 11;
    private int m_rowNum = 0;

    //
    // The cell styles for each of the base columns in the spreadsheet.
    private HSSFCellStyle[] m_CellStyles;

    //
    // Column widths

/*    private static final int CW_MESSAGE = 3144;
    private static final int CW_USA = 1060;
    private static final int CW_ITEM_DESC = 13130;
    private static final int CW_REG_BASE = 2304;
    private static final int CW_PROMO_COST = 2304;

    private static final int CW_STOCK_PACK = 1608;
    private static final int CW_C_MARKET = 2304;

    private static final int CW_PKG = 1206;
    private static final int CW_ITEM_NO = 1828;
    private static final int CW_ORD_QTY = 1610;
    private static final int CW_ITEM_NO2 = 2230;
    private static final int CW_ORD_QTY2 = 2230;*/

    private static final int CW_MESSAGE = 3300;
    private static final int CW_ITEM_DESC = 11500;
    private static final int CW_REG_BASE = 2304;
    private static final int CW_PROMO_COST = 2304;
    private static final int CW_STOCK_PACK = 1608;
    private static final int CW_UPC = 2500;
    private static final int CW_EMERY_ITEM_NO = 2304;
    private static final int CW_ITEM_NO = 1830;
    private static final int CW_ORD_QTY = 1610;
    private static final int CW_ITEM_NO2 = 2600;
    private static final int CW_ORD_QTY2 = 1610;

    /**
     * Default constructor. Initialize report variables.
     */
    public ShowPreprint() {
        super();

        ClassLoader classloader =
                org.apache.poi.poifs.filesystem.POIFSFileSystem.class.getClassLoader();
        URL res = classloader.getResource(
                "org/apache/poi/poifs/filesystem/POIFSFileSystem.class");
        String path = res.getPath();
        System.out.println("POI Core came from " + path);

        classloader = org.apache.poi.POIXMLDocument.class.getClassLoader();
        res = classloader.getResource("org/apache/poi/POIXMLDocument.class");
        path = res.getPath();
        System.out.println("POI OOXML came from " + path);

        m_Wrkbk = new HSSFWorkbook();
        m_Sheet = m_Wrkbk.createSheet();
        setupWorkbook();

        m_ShowDesc = "";
        m_PromoId = "";
        m_PacketId = "";
        m_CustomerId = null;
        m_StmtShowData = null;
        m_StmtShowVendors = null;
        m_StmtShowPromoHdr = null;
        m_StmtQtBuys = null;
        m_StmtShowPacketHdr = null;


    }

    /**
     * Cleanup any allocated resources.
     */
    @Override
    public void finalize() throws Throwable {
        m_Sheet = null;
        m_Wrkbk = null;
        m_ShowDesc = null;
        m_PromoId = null;
        m_ImagePath = null;
        m_PacketId = null;

        super.finalize();
    }

    /**
     * Executes the queries and builds the output file
     *
     * @throws java.io.FileNotFoundException
     */
    private boolean buildOutputFile() throws FileNotFoundException {
        DecimalFormat df = new DecimalFormat("'$'0.00");
        HSSFRow row = null;
        int colCnt = BASE_COLS;
        ResultSet rsetShowVendors = null;
        ResultSet rsetShowData = null;
        ResultSet rsetQtyBuys = null;
        boolean result = false;
        String nbc = "";
        String prevFlc = "";
        String newFlc = "";
        String vendorId;
        String vendorName;
        String boothNo;
        String itemId;
        String itemDesc = null;
        String[] itemDescs = null;
        int qbCnt = 0;
        String desc1 = "";
        String desc2 = "";
        String promoId = null;
        String packTitle = "";

        try {
            m_FilePath = System.getProperty("showrpt.dir", "/mnt/promos/show_order_forms/");
            m_ImagePath = System.getProperty("showrptimg.dir", "/mnt/promos/show_order_forms/logo/");
            //m_FilePath = "reports/";
            //m_ImagePath = "images/";

            if (m_PacketId != null && !m_PacketId.equals("")) {
                m_StmtShowVendors.setString(1, m_PacketId);
            } else
                m_StmtShowVendors.setString(1, m_PromoId);

            m_StmtShowVendors.setString(2, m_ShowDesc);
            rsetShowVendors = m_StmtShowVendors.executeQuery();

            while (rsetShowVendors.next() && m_Status == RptServer.RUNNING) {
                try {
                    initReport();

                    vendorId = rsetShowVendors.getString("vendor_id");
                    vendorName = rsetShowVendors.getString("vendor_name");
                    boothNo = rsetShowVendors.getString("booth");
                    boothNo = boothNo == null ? "" : boothNo;
                    packTitle = rsetShowVendors.getString("pack_title");

                    if (m_PacketId == null || m_PacketId.equals(""))
                        promoId = m_PromoId;

                    setCurAction("Create show order form for vendor " + vendorId);

                    m_rowNum = createReportTitle(m_rowNum, packTitle);

                    if (m_PacketId != null && !m_PacketId.equals(""))
                        m_rowNum = createPacketHeader(m_PacketId, vendorName, boothNo, m_rowNum + 3);
                    else
                        m_rowNum = createPromoHeader(promoId, vendorName, boothNo, m_rowNum + 3);

                    m_rowNum = createRowCaptions(m_rowNum);

                    if (m_PacketId != null && !m_PacketId.equals(""))
                        m_StmtShowData.setString(1, m_PacketId);
                    else
                        m_StmtShowData.setString(1, promoId);

                    m_StmtShowData.setString(2, vendorId);
                    m_StmtShowData.setString(3, m_ShowDesc);
                    rsetShowData = m_StmtShowData.executeQuery();

                    while (rsetShowData.next()) {
                        itemId = rsetShowData.getString("item_id");
                        newFlc = rsetShowData.getString("flc_id");

                        if (newFlc != null && !newFlc.equals(prevFlc)) {
                            row = createFLCRow(m_rowNum, rsetShowData.getString("flc_desc"));
                            m_rowNum++;
                        }

                        prevFlc = newFlc;
                        row = createRow(m_rowNum, colCnt);
                        itemDesc = rsetShowData.getString("item_desc");

                        //
                        //wrap the item desc to two lines
                        itemDescs = wrapText(itemDesc, 80);

                        for (int i = 0; i < itemDescs.length; i++) {
                            switch (i) {
                                case 0:
                                    desc1 = itemDescs[i];
                                case 1:
                                    desc2 = itemDescs[i];
                            }
                        }

                        if (itemDescs.length > 1) {
                            row.setHeightInPoints((2 * m_Sheet.getDefaultRowHeightInPoints()));
                            row.getCell(1).setCellValue(new HSSFRichTextString(desc1 + "\n" + desc2));
                        } else
                            row.getCell(1).setCellValue(new HSSFRichTextString(itemDesc));

                        row.getCell(0).setCellValue(new HSSFRichTextString(rsetShowData.getString("message")));

                        if ((rsetShowData.getDate("dsb_date") != null && rsetShowData.getDate("dsb_date").compareTo(rsetShowData.getDate("sysdt")) >= 0) &&
                                rsetShowData.getDouble("base_cost") != rsetShowData.getDouble("future_base_cost"))
                            row.getCell(2).setCellValue(rsetShowData.getDouble("future_base_cost"));
                        else
                            row.getCell(2).setCellValue(rsetShowData.getDouble("base_cost"));

                        //create string buffers for price and qty so we can store all the QBs in the same row - 08/22/14
                        StringBuffer priceBuffer = new StringBuffer();
                        StringBuffer qtyBuffer = new StringBuffer();

                        //row.getCell(4).setCellValue(rsetShowData.getDouble("promo_base"));
                        priceBuffer.append(df.format(rsetShowData.getDouble("promo_base")));

                        nbc = rsetShowData.getString("nbc");
                        nbc = nbc == null ? "" : nbc;

                        qtyBuffer.append(new HSSFRichTextString(rsetShowData.getString("stock_pack") + nbc));

                        row.getCell(5).setCellValue(rsetShowData.getString("upc_code"));

                        row.getCell(6).setCellValue(new HSSFRichTextString(itemId));
                        row.getCell(7).setCellValue(new HSSFRichTextString(rsetShowData.getString("ace_sku")));
                        row.getCell(9).setCellValue(new HSSFRichTextString(rsetShowData.getString("ace_sku")));

                        //
                        //Qty Buy prices
                        try {
                            qbCnt = 0;

                            if (m_PacketId != null && !m_PacketId.equals(""))
                                m_StmtQtBuys.setString(1, m_PacketId != null ? m_PacketId : "");
                            else
                                m_StmtQtBuys.setString(1, promoId != null ? promoId : "");

                            m_StmtQtBuys.setString(2, itemId != null ? itemId : "");
                            rsetQtyBuys = m_StmtQtBuys.executeQuery();

                            //ebronwewell: merged QBs into one row and removed extra rows per Michael. - 08/22/14
                            while (rsetQtyBuys.next()) {
                                //
                                //We need qty buys upto second level only
                                if (qbCnt == 2)
                                    break;

                        /* unnecessary as of update on 08/22/14
                        m_rowNum = m_rowNum + 1;
                        row = createRow(m_rowNum, colCnt);
                        row.getCell(4).setCellValue(rsetQtyBuys.getDouble("price"));//Promo Cost
                        row.getCell(5).setCellValue(rsetQtyBuys.getInt("qty"));//Min Qty Reqd.
                        */

                                //append a new line - 08/22/14
                                priceBuffer.append("\n");
                                qtyBuffer.append("\n");

                                //append the price and qty to their respective buffer - 08/22/14
                                priceBuffer.append(df.format(rsetQtyBuys.getDouble("price")));
                                qtyBuffer.append(rsetQtyBuys.getInt("qty"));

                                qbCnt++;
                            }

                            //set cell value to the stringbuffer that holds the QB entries - 08/22/14
                            row.getCell(3).setCellValue(priceBuffer.toString());
                            row.getCell(4).setCellValue(qtyBuffer.toString());
                        } catch (Exception e) {
                            log.fatal("[ShowPreprint]", e);
                        } finally {
                            closeRSet(rsetQtyBuys);
                            rsetQtyBuys = null;
                            qbCnt = 0;
                        }

                        m_rowNum++;
                    }

                    createFootNote(m_rowNum);
                    createFooter(vendorName);

                    //
                    //Create separate report for each vendor
                    if (m_PacketId != null && !m_PacketId.equals(""))
                        generateReport(vendorName, m_PacketId);
                    else
                        generateReport(vendorName, promoId);
                } catch (Exception e) {
                    log.fatal("[ShowPreprint]", e);
                } finally {
                    closeRSet(rsetShowData);
                    rsetShowData = null;
                    itemDescs = null;
                    row = null;
                }
            }

            result = true;
        } catch (Exception ex) {
            m_ErrMsg.append("Your report had the following errors: \r\n");
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());

            log.fatal("[ShowPreprint]", ex);
        } finally {
            closeRSet(rsetShowData);
            rsetShowData = null;
            row = null;

            m_ImagePath = null;
        }

        return result;
    }


    private int createPromoHeader(String promoId, String vendName, String booth, int rowNum) {
        ResultSet rsetShowPromoHdr = null;
        String header;
        Format formatter = null;
        Date diaDate;
        Date shipDate;
        String dia_date;
        String ship_date;
        String terms;
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFFont fontVend_Booth;
        HSSFFont fontPromoHdr;
        HSSFFont fontPO_Booth;
        HSSFCellStyle styleVend_Booth;
        HSSFCellStyle stylePromoHdr;
        HSSFCellStyle stylePO_Booth;
        String hdrTxt1 = "";
        String hdrTxt2 = "";
        String packetId = "";

        try {
            formatter = new SimpleDateFormat("MM/dd/yyyy");
            m_StmtShowPromoHdr.setString(1, promoId);
            rsetShowPromoHdr = m_StmtShowPromoHdr.executeQuery();
            if (rsetShowPromoHdr.next()) {
                header = rsetShowPromoHdr.getString("header");
                header = header == null ? "" : header;

                if (header.indexOf(".", 0) != -1) {
                    hdrTxt1 = header.substring(0, header.indexOf(".") + 1);
                    hdrTxt2 = header.substring(header.indexOf(".") + 1);
                }

                packetId = rsetShowPromoHdr.getString("packet_id");
                packetId = packetId == null ? "" : packetId;

                //
                //This is needed to display on the footer.
                m_PacketFooter = packetId;

                diaDate = rsetShowPromoHdr.getDate("dia_date");
                if (diaDate != null)
                    dia_date = formatter.format(diaDate);
                else
                    dia_date = "";

                shipDate = rsetShowPromoHdr.getDate("ship_date");
                if (shipDate != null)
                    ship_date = formatter.format(shipDate);
                else
                    ship_date = "";

                terms = rsetShowPromoHdr.getString("terms_name");
                terms = terms == null ? "" : terms;

                //
                //Style for vendor and booth information
                fontVend_Booth = m_Wrkbk.createFont();
                fontVend_Booth.setFontHeightInPoints((short) 10);
                fontVend_Booth.setFontName("Arial");
                fontVend_Booth.setBold(true);

                styleVend_Booth = m_Wrkbk.createCellStyle();
                styleVend_Booth.setFont(fontVend_Booth);
                styleVend_Booth.setAlignment(HorizontalAlignment.LEFT);


                //
                //Style for promo header info
                fontPromoHdr = m_Wrkbk.createFont();
                fontPromoHdr.setFontHeightInPoints((short) 7);
                fontPromoHdr.setFontName("Arial");
                fontPromoHdr.setBold(false);

                //
                //Style for po# and booth#
                fontPO_Booth = m_Wrkbk.createFont();
                fontPO_Booth.setFontHeightInPoints((short) 9);
                fontPO_Booth.setFontName("Arial");
                fontPO_Booth.setBold(true);

                stylePromoHdr = m_Wrkbk.createCellStyle();
                stylePromoHdr.setFont(fontPromoHdr);
                stylePromoHdr.setAlignment(HorizontalAlignment.LEFT);

                stylePO_Booth = m_Wrkbk.createCellStyle();
                stylePO_Booth.setFont(fontPO_Booth);
                stylePO_Booth.setAlignment(HorizontalAlignment.LEFT);

                //
                //Style for Requested Ship date
                HSSFFont fontShipDate = m_Wrkbk.createFont();
                fontShipDate.setFontHeightInPoints((short) 9);
                fontShipDate.setFontName("Arial");
                fontShipDate.setBold(true);

                HSSFCellStyle styleShipDate = m_Wrkbk.createCellStyle();
                styleShipDate.setFont(fontShipDate);
                styleShipDate.setAlignment(HorizontalAlignment.LEFT);

                row = m_Sheet.createRow(rowNum);

                cell = row.createCell(6);
                cell.setCellStyle(stylePromoHdr);
                cell.setCellValue(new HSSFRichTextString("Packet: " + packetId));


                cell = row.createCell(9); // ebrownewell - 01/12/15
                cell.setCellStyle(stylePO_Booth);
                cell.setCellValue(new HSSFRichTextString("Cust #: ___________"));

                rowNum++;
                row = m_Sheet.createRow(rowNum);

                //Vendor Name
                cell = row.createCell(0);
                cell.setCellStyle(styleVend_Booth);
                cell.setCellValue(new HSSFRichTextString(vendName));

                //Promo Order Deadline
                cell = row.createCell(6);
                cell.setCellStyle(stylePromoHdr);
                cell.setCellValue(new HSSFRichTextString("Order Deadline: " + dia_date));

                //PO #
                cell = row.createCell(9);
                cell.setCellStyle(stylePO_Booth);
                cell.setCellValue(new HSSFRichTextString("PO #:__________"));

                rowNum++;
                row = m_Sheet.createRow(rowNum);

                //Booth #
                cell = row.createCell(0);
                cell.setCellStyle(styleVend_Booth);
                cell.setCellValue(new HSSFRichTextString("Booth #: " + booth));

                //Promo header txt1
                cell = row.createCell(2);
                cell.setCellStyle(stylePromoHdr);
                cell.setCellValue(new HSSFRichTextString(hdrTxt1));

                //Ship date
                cell = row.createCell(6);
                cell.setCellStyle(stylePromoHdr);
                cell.setCellValue(new HSSFRichTextString("Approx. Ship Date: " + ship_date));

                //Ship date
                cell = row.createCell(9);
                cell.setCellStyle(styleShipDate);
                cell.setCellValue(new HSSFRichTextString("Desired Ship Date"));

                rowNum++;
                row = m_Sheet.createRow(rowNum);

                //Promo Hdr text2
                cell = row.createCell(2);
                cell.setCellStyle(stylePromoHdr);
                cell.setCellValue(new HSSFRichTextString(hdrTxt2));

                //Terms
                cell = row.createCell(6);
                cell.setCellStyle(stylePromoHdr);
                cell.setCellValue(new HSSFRichTextString("Due: " + terms));

                //Ship date line
                cell = row.createCell(9);
                cell.setCellStyle(styleShipDate);
                cell.setCellValue(new HSSFRichTextString("________________"));
            }

        } catch (Exception ex) {
            log.fatal("[ShowPreprint]", ex);
        } finally {
            closeRSet(rsetShowPromoHdr);
            rsetShowPromoHdr = null;
        }

        return rowNum;

    }

    private int createPacketHeader(String packetId, String vendName, String booth, int rowNum) {
        ResultSet rsetShowPacketHdr = null;
        String header;
        Format formatter = null;
        Date diaDate;
        Date shipDate;
        String dia_date = "";
        String ship_date = "";
        String terms = "";
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFFont fontVend_Booth;
        HSSFFont fontPromoHdr;
        HSSFFont fontPO_Booth;
        HSSFCellStyle styleVend_Booth;
        HSSFCellStyle stylePromoHdr;
        HSSFCellStyle stylePO_Booth;
        String hdrTxt1 = "";
        String hdrTxt2 = "";

        try {
            formatter = new SimpleDateFormat("MM/dd/yyyy");
            m_StmtShowPacketHdr.setString(1, packetId);
            rsetShowPacketHdr = m_StmtShowPacketHdr.executeQuery();
            if (rsetShowPacketHdr.next()) {
                header = rsetShowPacketHdr.getString("header");
                header = header == null ? "" : header;

                if (header.indexOf(".", 0) != -1) {
                    hdrTxt1 = header.substring(0, header.indexOf(".") + 1);
                    hdrTxt2 = header.substring(header.indexOf(".") + 1);
                }

                packetId = rsetShowPacketHdr.getString("packet_id");
                packetId = packetId == null ? "" : packetId;

                //
                //This is needed to display on the footer.
                m_PacketFooter = packetId;

                diaDate = rsetShowPacketHdr.getDate("dia_date");
                if (diaDate != null)
                    dia_date = formatter.format(diaDate);
                else
                    dia_date = "";

                shipDate = rsetShowPacketHdr.getDate("ship_date");
                if (shipDate != null)
                    ship_date = formatter.format(shipDate);
                else
                    ship_date = "";

                terms = rsetShowPacketHdr.getString("terms");
                terms = terms == null ? "" : terms;
            }

            //
            //Style for vendor and booth information
            fontVend_Booth = m_Wrkbk.createFont();
            fontVend_Booth.setFontHeightInPoints((short) 10);
            fontVend_Booth.setFontName("Arial");
            fontVend_Booth.setBold(true);

            styleVend_Booth = m_Wrkbk.createCellStyle();
            styleVend_Booth.setFont(fontVend_Booth);
            styleVend_Booth.setAlignment(HorizontalAlignment.LEFT);

            //
            //Style for promo header info
            fontPromoHdr = m_Wrkbk.createFont();
            fontPromoHdr.setFontHeightInPoints((short) 7);
            fontPromoHdr.setFontName("Arial");
            fontPromoHdr.setBold(false);

            //
            //Style for po# and booth#
            fontPO_Booth = m_Wrkbk.createFont();
            fontPO_Booth.setFontHeightInPoints((short) 9);
            fontPO_Booth.setFontName("Arial");
            fontPO_Booth.setBold(true);

            stylePromoHdr = m_Wrkbk.createCellStyle();
            stylePromoHdr.setFont(fontPromoHdr);
            stylePromoHdr.setAlignment(HorizontalAlignment.LEFT);

            stylePO_Booth = m_Wrkbk.createCellStyle();
            stylePO_Booth.setFont(fontPO_Booth);
            stylePO_Booth.setAlignment(HorizontalAlignment.LEFT);

            //
            //Style for Requested Ship date
            HSSFFont fontShipDate = m_Wrkbk.createFont();
            fontShipDate.setFontHeightInPoints((short) 9);
            fontShipDate.setFontName("Arial");
            fontShipDate.setBold(true);

            HSSFCellStyle styleShipDate = m_Wrkbk.createCellStyle();
            styleShipDate.setFont(fontShipDate);
            styleShipDate.setAlignment(HorizontalAlignment.LEFT);

            row = m_Sheet.createRow(rowNum);

            cell = row.createCell(6);
            cell.setCellStyle(stylePromoHdr);
            cell.setCellValue(new HSSFRichTextString("Packet: " + packetId));


            cell = row.createCell(9); // ebrownewell - 01/12/15
            cell.setCellStyle(stylePO_Booth);
            cell.setCellValue(new HSSFRichTextString("Cust #: ___________"));

            rowNum++;
            row = m_Sheet.createRow(rowNum);

            //Vendor Name
            cell = row.createCell(0);
            cell.setCellStyle(styleVend_Booth);
            cell.setCellValue(new HSSFRichTextString(vendName));

            //Promo Order Deadline
            cell = row.createCell(6);
            cell.setCellStyle(stylePromoHdr);
            cell.setCellValue(new HSSFRichTextString("Order Deadline: " + dia_date));

            //PO #
            cell = row.createCell(9);
            cell.setCellStyle(stylePO_Booth);
            cell.setCellValue(new HSSFRichTextString("PO #:__________"));

            rowNum++;
            row = m_Sheet.createRow(rowNum);

            //Booth #
            cell = row.createCell(0);
            cell.setCellStyle(styleVend_Booth);
            cell.setCellValue(new HSSFRichTextString("Booth #: " + booth));

            //Promo header txt1
            cell = row.createCell(2);
            cell.setCellStyle(stylePromoHdr);
            cell.setCellValue(new HSSFRichTextString(hdrTxt1));

            //Ship date
            cell = row.createCell(6);
            cell.setCellStyle(stylePromoHdr);
            cell.setCellValue(new HSSFRichTextString("Approx. Ship Date: " + ship_date));

            //Ship date
            cell = row.createCell(9);
            cell.setCellStyle(styleShipDate);
            cell.setCellValue(new HSSFRichTextString("Desired Ship Date"));

            rowNum++;
            row = m_Sheet.createRow(rowNum);

            //Promo Hdr text2
            cell = row.createCell(2);
            cell.setCellStyle(stylePromoHdr);
            cell.setCellValue(new HSSFRichTextString(hdrTxt2));

            //Terms
            cell = row.createCell(6);
            cell.setCellStyle(stylePromoHdr);
            cell.setCellValue(new HSSFRichTextString("Due: " + terms));

            //Ship date line
            cell = row.createCell(9);
            cell.setCellStyle(styleShipDate);
            cell.setCellValue(new HSSFRichTextString("________________"));

        } catch (Exception ex) {
            log.fatal("[ShowPreprint]", ex);
        } finally {
            closeRSet(rsetShowPacketHdr);
            rsetShowPacketHdr = null;
        }

        return rowNum;
    }

    private void initReport() {
        m_Wrkbk = new HSSFWorkbook();
        m_Sheet = m_Wrkbk.createSheet();
        m_Sheet.setRepeatingRows(CellRangeAddress.valueOf("1:9"));

        //set margins so the sheet will look right when printed
        m_Sheet.setMargin(Sheet.RightMargin, 0);
        m_Sheet.setMargin(Sheet.LeftMargin, 0.25);
        m_Sheet.setMargin(Sheet.BottomMargin, 0.75);
        m_Sheet.setMargin(Sheet.TopMargin, 0.3);

        m_Sheet.getPrintSetup().setFooterMargin(0);

        m_Sheet.setHorizontallyCenter(true);

        //m_Sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.A4_PAPERSIZE);

        m_Sheet.getPrintSetup().setLandscape(true);
        m_rowNum = 0;

        setupWorkbook();
    }

    /**
     * Runs the report and creates any output that is needed.
     */
    private void generateReport(String vendorName, String promoId) {
        StringBuffer fileName = new StringBuffer();
        FileOutputStream OutFile = null;
        String tmp = null;

        try {
            //
            // Build the report file name
            tmp = Long.toString(System.currentTimeMillis());
            //
            //Make sure vendor name doesn't contain any backlash character.
            vendorName = backlashReplace(vendorName);
            fileName.append("showorderfrm");
            fileName.append("-");
            fileName.append(vendorName);
            fileName.append("(").append(promoId).append(")");
            fileName.append("-");
            fileName.append(tmp.substring(tmp.length() - 5));
            fileName.append(".xls");

            OutFile = new FileOutputStream(m_FilePath + fileName, false);

            m_Wrkbk.write(OutFile);

            try {
                OutFile.close();
            } catch (Exception e) {
                log.error("[ShowPreprint]", e);
            }
        } catch (Exception ex) {
            log.error("[ShowPreprint]:", ex);
            m_ErrMsg.append("The report had the following Error: \r\n");
            m_ErrMsg.append(ex.getClass().getName()).append("\r\n").append(ex.getMessage());
        } finally {
            m_Sheet = null;
            m_Wrkbk = null;
        }
    }

    /**
     * Closes all the sql statements so they release the db cursors.
     */
    private void closeStatements() {
        closeStmt(m_StmtShowVendors);
        closeStmt(m_StmtShowPromoHdr);
        closeStmt(m_StmtShowData);
        closeStmt(m_StmtQtBuys);
        closeStmt(m_StmtShowPacketHdr);
    }

    protected HSSFCell createCaptionCell(HSSFRow row, int col, String caption, HSSFCellStyle stylCaptions) {
        HSSFCell cell = null;
        HSSFCellStyle m_CSCaption = null;
        HSSFFont font = null;

        if (row != null) {
            font = m_Wrkbk.createFont();
            font.setFontHeightInPoints((short) 8);
            font.setFontName("Arial");
            font.setBold(true);

            m_CSCaption = m_Wrkbk.createCellStyle();
            m_CSCaption.setFont(font);
            m_CSCaption.setAlignment(HorizontalAlignment.CENTER);
            m_CSCaption.setWrapText(true);

            //
            //Shading
            m_CSCaption.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            m_CSCaption.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            cell = row.createCell(col);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(stylCaptions);
            cell.setCellValue(new HSSFRichTextString(caption != null ? caption : ""));
        }

        return cell;
    }

    /**
     */
    private void createFootNote(int rowNum) {
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFFont fontFootNote;
        HSSFCellStyle styleFootNote;

        String ftNoteExtraEntry = "EEE=Earn Extra Entry";
        String ftNoteHotbuy = "HB=Hot Buy";
        String ftNoteIncredibleDeal = "ID=Incredible Deal";
        String ftNoteNoBrkCarton = "N=No Broken Cartons";
        String ftNoteSlamDunk = "SD=Slam Dunk - Deep Discounts offered to show attendees only";

        //
        //Style for Foot note
        fontFootNote = m_Wrkbk.createFont();
        fontFootNote.setFontHeightInPoints((short) 7);
        fontFootNote.setFontName("Arial");
        fontFootNote.setBold(false);

        styleFootNote = m_Wrkbk.createCellStyle();
        styleFootNote.setFont(fontFootNote);
        styleFootNote.setAlignment(HorizontalAlignment.LEFT);

        //
        //Extra Entry
        if (m_ExtraEntry) {
            rowNum = rowNum + 1;
            row = m_Sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(styleFootNote);
            cell.setCellValue(new HSSFRichTextString(ftNoteExtraEntry));
        }

        //
        //Hot Buy
        if (m_Hotbuy) {
            rowNum = rowNum + 1;
            row = m_Sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(styleFootNote);
            cell.setCellValue(new HSSFRichTextString(ftNoteHotbuy));
        }

        //
        //Incredible Deal
        if (m_Incredible) {
            rowNum = rowNum + 1;
            row = m_Sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(styleFootNote);
            cell.setCellValue(new HSSFRichTextString(ftNoteIncredibleDeal));
        }

        //
        //No Broken Cartons
        if (m_NoBrkCarton) {
            rowNum = rowNum + 1;
            row = m_Sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(styleFootNote);
            cell.setCellValue(new HSSFRichTextString(ftNoteNoBrkCarton));
        }

        //
        //Slam Dunk
        if (m_SlamDunk) {
            rowNum = rowNum + 1;
            row = m_Sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(styleFootNote);
            cell.setCellValue(new HSSFRichTextString(ftNoteSlamDunk));
        }

    }

    /**
     */
    private void createFooter(String vendName) {
        HSSFFooter footer = m_Sheet.getFooter();

        footer.font("Arial", "Regular");
        footer.fontSize((short) 8);

        //
        //Vendor Name
        footer.setLeft(vendName);

        //
        //Page numbers
        footer.setCenter("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        //
        //Packet
        //added Vendor and page number to right side - 08/22/14
        if (m_PacketId != null && !m_PacketId.equals(""))
            footer.setRight("Packet: " + m_PacketId + "\n" + HeaderFooter.page() + " of " + HeaderFooter.numPages());
        else
            footer.setRight("Packet: " + m_PacketFooter + "\n" + HeaderFooter.page() + " of " + HeaderFooter.numPages());
    }

    /**
     */
    public int createReportTitle(int rowNum, String packetTitle) {
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFFont fontShowTitle;
        HSSFFont fontEmeryLogo;
        HSSFCellStyle styleShowTitle;
        HSSFCellStyle styleEmeryLogo;
        InputStream is = null;
        CreationHelper helper = null;
        Drawing drawing = null;
        ClientAnchor anchor = null;
        Picture pict = null;
        int pictureIdx = 0;
        byte[] imgBytes = null;

        row = m_Sheet.createRow(rowNum);
        row.setHeightInPoints((float) 37.50);

        //
        //Style for Show Title
        fontShowTitle = m_Wrkbk.createFont();
        fontShowTitle.setFontHeightInPoints((short) 14);
        fontShowTitle.setFontName("Arial");
        fontShowTitle.setBold(true);
        fontShowTitle.setItalic(true);

        styleShowTitle = m_Wrkbk.createCellStyle();
        styleShowTitle.setFont(fontShowTitle);
        styleShowTitle.setAlignment(HorizontalAlignment.LEFT);
        styleShowTitle.setVerticalAlignment(VerticalAlignment.TOP);
        styleShowTitle.setFillBackgroundColor(new HSSFColor.RED().getIndex());

        cell = row.createCell(0);
        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(styleShowTitle);
        cell.setCellValue(new HSSFRichTextString(packetTitle));

        //
        //Style for Emery logo
        fontEmeryLogo = m_Wrkbk.createFont();
        fontEmeryLogo.setFontHeightInPoints((short) 12);
        fontEmeryLogo.setFontName("Arial");
        fontEmeryLogo.setBold(true);

        styleEmeryLogo = m_Wrkbk.createCellStyle();
        styleEmeryLogo.setFont(fontEmeryLogo);
        styleEmeryLogo.setAlignment(HorizontalAlignment.LEFT);
        styleShowTitle.setFillForegroundColor(new HSSFColor.RED().getIndex());
        styleShowTitle.setFillBackgroundColor(new HSSFColor.RED().getIndex());

        //
        //Style for po# and booth#
        HSSFFont fontPO_Booth = m_Wrkbk.createFont();
        fontPO_Booth.setFontHeightInPoints((short) 9);
        fontPO_Booth.setFontName("Arial");
        fontPO_Booth.setBold(true);

        HSSFCellStyle stylePO_Booth = m_Wrkbk.createCellStyle();
        stylePO_Booth.setFont(fontPO_Booth);
        stylePO_Booth.setAlignment(HorizontalAlignment.LEFT);

        try {
            //is = new FileInputStream(m_ImagePath + "old_logo.jpg");
            is = new FileInputStream(m_ImagePath + "logo.jpg");

            imgBytes = IOUtils.toByteArray(is);
            pictureIdx = m_Wrkbk.addPicture(imgBytes, HSSFWorkbook.PICTURE_TYPE_JPEG);

            helper = m_Wrkbk.getCreationHelper();
            drawing = m_Sheet.createDrawingPatriarch();
            anchor = helper.createClientAnchor();
            anchor.setCol1(6);
            anchor.setRow1(rowNum);
            pict = drawing.createPicture(anchor, pictureIdx);

            //
            //auto-size picture
            pict.resize(3,2);
        } catch (Exception ex) {
            log.error("[ShowPreprint]", ex);
        } finally {
            //
            // Close the inputstream
            if (is != null) {
                try {
                    is.close();
                } catch (Exception e) {
                    //
                }
            }
        }

        return ++rowNum;
    }

    /**
     * Creates the captions for the vendor filter.
     *
     * @see SubRpt#createCaptions(int rowNum)
     */
    public int createRowCaptions(int rowNum) {
        HSSFRow row = null;
        HSSFCellStyle styleCaptionsRow = null;
        HSSFFont fontCaptionsRow = null;
        int col = 0;

        fontCaptionsRow = m_Wrkbk.createFont();
        fontCaptionsRow.setFontHeightInPoints((short) 8);
        fontCaptionsRow.setFontName("Arial");
        fontCaptionsRow.setBold(true);

        styleCaptionsRow = m_Wrkbk.createCellStyle();
        styleCaptionsRow.setFont(fontCaptionsRow);
        styleCaptionsRow.setAlignment(HorizontalAlignment.CENTER);

        //
        //Shading
        styleCaptionsRow.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        styleCaptionsRow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //
        //Border
        styleCaptionsRow.setBorderTop(BorderStyle.THIN);
        styleCaptionsRow.setBorderBottom(BorderStyle.THIN);
        styleCaptionsRow.setBorderLeft(BorderStyle.THIN);
        styleCaptionsRow.setBorderRight(BorderStyle.THIN);

        styleCaptionsRow.setWrapText(true);

        //
        //Additional row for QB1
        rowNum++;
        row = m_Sheet.createRow(rowNum);
        row.setRowStyle(styleCaptionsRow);

        createCaptionCell(row, col, "Message", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_MESSAGE);
        createCaptionCell(row, col, "Item Description", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_ITEM_DESC);
        createCaptionCell(row, col, "Reg Base", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_REG_BASE);
        createCaptionCell(row, col, "Promo Cost\nQB1\nQB2", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_PROMO_COST);
        createCaptionCell(row, col, "Stk Pk\nQB1\nQB2", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_STOCK_PACK);
        createCaptionCell(row, col, "UPC", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_UPC);
        createCaptionCell(row, col, "Emery #", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_EMERY_ITEM_NO);
        createCaptionCell(row, col, "Item #", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_ITEM_NO);
        createCaptionCell(row, col, "Ord Qty", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_ORD_QTY);
        createCaptionCell(row, col, "Item #", styleCaptionsRow);
        m_Sheet.setColumnWidth(col++, CW_ITEM_NO2);
        createCaptionCell(row, col, "Ord Qty", styleCaptionsRow);
        m_Sheet.setColumnWidth(col, CW_ORD_QTY2);

        return ++rowNum;
    }

    /**
     * Creates the show data export report.
     *
     * @see com.emerywaterhouse.rpt.server.Report#createReport()
     */
    @Override
    public boolean createReport() {
        boolean created = false;
        m_Status = RptServer.RUNNING;

        try {
            m_EdbConn = m_RptProc.getEdbConn();
            //m_EdbConn = getConnectionPg();

            if (prepareStatements())
                created = buildOutputFile();
        } catch (Exception ex) {
            log.fatal("[ShowPreprint]", ex);
        } finally {
            closeStatements();

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
    private HSSFRow createRow(int rowNum, int colCnt) {
        HSSFRow row = null;
        HSSFCell cell = null;

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
     * Creates a row in the worksheet.
     *
     * @param rowNum The row number.
     * @return The formatted row of the spreadsheet.
     */
    private HSSFRow createFLCRow(int rowNum, String flcDesc) {
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFCellStyle styleTitle1;   // Bold, centered
        HSSFFont fontTitle;   // Bold, centered

        if (m_Sheet == null)
            return row;

        fontTitle = m_Wrkbk.createFont();
        fontTitle.setFontHeightInPoints((short) 9);
        fontTitle.setFontName("Arial");
        fontTitle.setBold(true);
        fontTitle.setColor(HSSFColor.WHITE.index);

        styleTitle1 = m_Wrkbk.createCellStyle();
        styleTitle1.setFont(fontTitle);
        styleTitle1.setAlignment(HorizontalAlignment.LEFT);

        //
        //Shading
        styleTitle1.setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);
        styleTitle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //
        //Assign border for each cell of the row
        styleTitle1.setBorderLeft(BorderStyle.THIN);
        styleTitle1.setBorderRight(BorderStyle.THIN);

        row = m_Sheet.createRow(rowNum);
        row.setRowStyle(styleTitle1);

        //
        // set the type and style of the cell.
        cell = row.createCell(0);
        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(styleTitle1);
        cell.setCellValue(new HSSFRichTextString(flcDesc));

        return row;
    }


    /**
     * Prepares the sql queries for execution.
     */
    private boolean prepareStatements() {
        StringBuffer sql = new StringBuffer();
        boolean isPrepared = false;

        if (m_EdbConn != null) {
            try {
                sql.append("select ");
                sql.append("   decode(vs.name, null, vendor.name, vs.name) as vendor_name, ");

                if (m_PacketId == null || m_PacketId.equals("")) {
                    sql.append("   promo_item.promo_id, ");
                }

                sql.append("   item_entity_attr.vendor_id,show_vendor.booth,packet.title as pack_title ");
                sql.append("from ");
                sql.append("   preprint_item ");
                sql.append("   join promo_item on promo_item.promo_item_id = preprint_item.promo_item_id ");
                sql.append("   join promotion on  promo_item.promo_id = promotion.promo_id ");
                sql.append("   join packet on packet.packet_id = promotion.packet_id ");

                if (m_PacketId != null && !m_PacketId.equals("")) {
                    sql.append(" and packet.packet_id = ? ");
                } else {
                    sql.append(" and  promotion.promo_id = ? ");
                }
                sql.append("   join item_entity_attr on promo_item.item_ea_id = item_entity_attr.item_ea_id  ");
                sql.append("   join vendor on item_entity_attr.vendor_id = vendor.vendor_id  ");
                sql.append("   left outer join vendor_shortname vs on item_entity_attr.vendor_id = vs.vendor_id ");
                sql.append("   left outer join show_vendor on item_entity_attr.vendor_id = show_vendor.vendor_id and ");
                sql.append("   show_vendor.show_id in(select show_id from show where show.name = ?) ");
                sql.append("group by ");
                sql.append("   vs.name,vendor.name, ");

                if (m_PacketId == null || m_PacketId.equals("")) {
                    sql.append("   promo_item.promo_id, ");
                }

                sql.append("   item_entity_attr.vendor_id,show_vendor.booth,packet.title ");
                sql.append("order by ");
                sql.append("   vendor.name ");
                m_StmtShowVendors = m_EdbConn.prepareStatement(sql.toString());

                sql.setLength(0);
                sql.append("select ");
                sql.append("   title,dia_date,dsb_date,dsa_date,ship_date,reorder_start,reorder_end,promo_type_id,packet_id, ");
                sql.append("   promotion.term_id,ejd.terms_procs.debabelize(terms.name) as terms_name,header,footer,ship_text ");
                sql.append("from ");
                sql.append("   promotion, preprint, terms ");
                sql.append("where ");
                sql.append("   promotion.promo_id = preprint.promo_id and ");
                sql.append("   promotion.term_id = terms.term_id and promotion.promo_id = ? ");
                m_StmtShowPromoHdr = m_EdbConn.prepareStatement(sql.toString());

                sql.setLength(0);
                sql.append("select ");
                sql.append("   vendor.name as vendor_name, promo_item.promo_id, promo_item.item_id, item_entity_attr.vendor_id,flc.description as flc_desc, ");
                sql.append("   item_entity_attr.description as item_desc, ");
                sql.append("   decode(broken_case.description, 'ALLOW BROKEN CASES', '', 'N') as nbc, ejd_item_warehouse.stock_pack, unit as ship_unit, ");
                sql.append("   ejd_item_price.sell as base_cost,  ");
                sql.append("   ejd_item_price.retail_c as reg_retail,  ");
                sql.append("   ejd_price_procs.get_sell_for_date(promo_item.item_ea_id,ejd_item_warehouse.warehouse_id, dsb_date) as future_base_cost,   ");
                sql.append("   ejd_price_procs.get_retailc_for_date(promo_item.item_ea_id, ejd_item_warehouse.warehouse_id, dsb_date) as future_reg_retail, ");
                sql.append("   promo_base, message, ");
                sql.append("   trunc(dsb_date) as dsb_date, trunc(now()) as sysdt, ejd_item.flc_id, show_vendor.booth,item_entity_attr.item_id as usa_item,  ");
                sql.append("   ace_item_xref.ace_sku, ejd_item_whs_upc.upc_code ");
                sql.append("from  ");
                sql.append("preprint_item ");
                sql.append("join promo_item on promo_item.promo_item_id = preprint_item.promo_item_id ");
                sql.append("join promotion on  promo_item.promo_id = promotion.promo_id ");
                sql.append("join packet on packet.packet_id = promotion.packet_id and ");

                if (m_PacketId != null && !m_PacketId.equals("")) {
                    sql.append(" packet.packet_id = ? ");
                } else {
                    sql.append(" promotion.promo_id = ? ");
                }

                sql.append("join item_entity_attr on promo_item.item_ea_id = item_entity_attr.item_ea_id ");
                sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
                sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = 1 ");
                sql.append("join ejd_item_whs_upc on ejd_item_warehouse.ejd_item_id = ejd_item_whs_upc.ejd_item_id and ejd_item_warehouse.warehouse_id = ejd_item_whs_upc.warehouse_id ");
                sql.append("and ejd_item_whs_upc.primary_upc = 1 ");
                sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id ");
                sql.append("join ship_unit on item_entity_attr.ship_unit_id = ship_unit.unit_id ");
                sql.append("join vendor on item_entity_attr.vendor_id = vendor.vendor_id and ");
                sql.append("item_entity_attr.vendor_id = ? ");
                sql.append("join broken_case on ejd_item.broken_case_id = broken_case.broken_case_id  ");
                sql.append("join  flc on ejd_item.flc_id = flc.flc_id ");
                sql.append("left outer join ejd_item_attribute on item_entity_attr.ejd_item_id = ejd_item_attribute.ejd_item_id and ");
                sql.append("ejd_item_attribute.attribute_value_id in ( select attribute_value_id from attribute a , attribute_value av where a.attribute_id = av.attribute_id and av.value = 'MADE IN USA') ");
                sql.append("left outer join bmi_item on bmi_item.item_id = item_entity_attr.item_id  ");
                sql.append("left outer join show_vendor on item_entity_attr.vendor_id = show_vendor.vendor_id and ");
                sql.append("show_vendor.show_id in(select show_id from show where show.name = ?) ");
                sql.append("left outer join ace_item_xref on item_entity_attr.item_id = ace_item_xref.item_id ");
                sql.append("group by ");
                sql.append("vendor.name,item_entity_attr.vendor_id,promo_item.promo_id,promo_item.item_id,flc.description,item_entity_attr.description,broken_case.description, ejd_item_warehouse.stock_pack,unit, ");
                sql.append("ejd_item_price.sell, ");
                sql.append("ejd_item_price.retail_c, ");
                sql.append("ejd_price_procs.get_sell_for_date(promo_item.item_ea_id,ejd_item_warehouse.warehouse_id, dsb_date), ");
                sql.append("ejd_price_procs.get_retailc_for_date(promo_item.item_ea_id,ejd_item_warehouse.warehouse_id, dsb_date),  ");
                sql.append("promo_base, item_entity_attr.vendor_id, message, bmi_item.web_descr, item_entity_attr.description,  ");
                sql.append("dsb_date, now(), ejd_item.flc_id, booth,   ");
                sql.append("flc.description,item_entity_attr.item_id,   ");
                sql.append("ace_item_xref.ace_sku, ejd_item_whs_upc.upc_code ");
                sql.append("order by flc_desc,item_desc ");
                m_StmtShowData = m_EdbConn.prepareStatement(sql.toString());

                sql.setLength(0);
                sql.append(" select quantity_buy_item.item_id, ");
                sql.append("  item_entity_attr.description, ");
                sql.append(" quantity_buy_item.min_qty qty, ");
                sql.append(" quantity_buy_item.discount_value price,promotion.promo_id ");
                sql.append(" from packet ");
                sql.append(" join promotion on promotion.packet_id = packet.packet_id and ");

                if (m_PacketId != null && !m_PacketId.equals("")) {
                    sql.append(" packet.packet_id = ? ");
                } else {
                    sql.append(" promotion.promo_id = ? ");
                }

                sql.append(" join promo_item on promo_item.promo_id = promotion.promo_id ");
                sql.append(" join quantity_buy on quantity_buy.packet_id = packet.packet_id ");
                sql.append(" join discount on discount.discount_id = quantity_buy.discount_id ");
                sql.append(" join quantity_buy_item on quantity_buy_item.qty_buy_id = quantity_buy.qty_buy_id and ");
                sql.append(" quantity_buy_item.item_ea_id = promo_item.item_ea_id ");
                sql.append(" join item_entity_attr on item_entity_attr.item_ea_id = quantity_buy_item.item_ea_id  ");
                sql.append(" where ");
                sql.append(" item_entity_attr.item_id = ? ");
                sql.append(" order by qty ");
                m_StmtQtBuys = m_EdbConn.prepareStatement(sql.toString());

                sql.setLength(0);
                sql.append("select ");
                sql.append("   packet_id,dia_date,ship_date,terms,header ");
                sql.append("from ");
                sql.append("   show_packet ");
                sql.append("where ");
                sql.append("   packet_id = ? ");
                m_StmtShowPacketHdr = m_EdbConn.prepareStatement(sql.toString());

                isPrepared = true;
            } catch (SQLException ex) {
                log.error("[ShowPreprint]", ex);
            } finally {
                sql = null;
            }
        } else
            log.error("[ShowPreprint] prepareStatements - null Database connection");

        return isPrepared;
    }

    /**
     * Sets the parameters of this report.
     *
     * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
     */
    @Override
    public void setParams(ArrayList<Param> params) {
        Param param = null;
        int pcount = params.size();

        for (int i = 0; i < pcount; i++) {
            param = params.get(i);

            if (param.name.equals("showdesc"))
                m_ShowDesc = param.value;

            if (param.name.equals("promoid"))
                m_PromoId = param.value;

            if (param.name.equals("packetid"))
                m_PacketId = param.value;

            if (param.name.equals("extraentry"))
                m_ExtraEntry = param.value.equalsIgnoreCase("true");

            if (param.name.equals("hotbuy"))
                m_Hotbuy = param.value.equalsIgnoreCase("true");

            if (param.name.equals("incredible"))
                m_Incredible = param.value.equalsIgnoreCase("true");

            if (param.name.equals("nobrkcarton"))
                m_NoBrkCarton = param.value.equalsIgnoreCase("true");

            if (param.name.equals("slamdunk"))
                m_SlamDunk = param.value.equalsIgnoreCase("true");

            if (param.name.equals("customerid")) {
                m_CustomerId = param.value;
            }

        }
    }

    /**
     * Sets up the styles for the cells based on the column data.  Does any other inititialization
     * needed by the workbook.
     */
    private void setupWorkbook() {
        HSSFCellStyle styleText;      // Text centered
        HSSFCellStyle styleItemDesc;  // Text with left alignment
        HSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
        HSSFCellStyle styleMoneyWrap; // Money ($#,##0.00_);[Red]($#,##0.00) with line wrap enabled - 08/22/14
        HSSFCellStyle styleTextWrap;  // Text Centered with line wrap enabled - 08/22/14

        HSSFFont font = m_Wrkbk.createFont();
        font.setFontHeightInPoints((short) 8);
        font.setFontName("Arial");
        font.setBold(true);

        styleText = m_Wrkbk.createCellStyle();
        styleText.setFont(font);
        styleText.setAlignment(HorizontalAlignment.CENTER);

        //
        //Assign border for each cell of the row
        styleText.setBorderTop(BorderStyle.THIN);
        styleText.setBorderBottom(BorderStyle.THIN);
        styleText.setBorderLeft(BorderStyle.THIN);
        styleText.setBorderRight(BorderStyle.THIN);

        //new style for Stock pack to include line wrap - 08/22/14
        styleTextWrap = m_Wrkbk.createCellStyle();
        styleTextWrap.setFont(font);
        styleTextWrap.setAlignment(HorizontalAlignment.CENTER);
        styleTextWrap.setWrapText(true);

        styleTextWrap.setBorderTop(BorderStyle.THIN);
        styleTextWrap.setBorderBottom(BorderStyle.THIN);
        styleTextWrap.setBorderLeft(BorderStyle.THIN);
        styleTextWrap.setBorderRight(BorderStyle.THIN);

        //
        //Style for item desc
        styleItemDesc = m_Wrkbk.createCellStyle();
        styleItemDesc.setFont(font);
        styleItemDesc.setWrapText(true);
        styleItemDesc.setAlignment(HorizontalAlignment.LEFT);

        //
        //Assign border for each cell of the row
        styleItemDesc.setBorderTop(BorderStyle.THIN);
        styleItemDesc.setBorderBottom(BorderStyle.THIN);
        styleItemDesc.setBorderLeft(BorderStyle.THIN);
        styleItemDesc.setBorderRight(BorderStyle.THIN);

        styleMoney = m_Wrkbk.createCellStyle();
        styleMoney.setFont(font);
        styleMoney.setAlignment(HorizontalAlignment.RIGHT);
        styleMoney.setDataFormat((short) 8);

        //
        //Assign border for each cell of the row
        styleMoney.setBorderTop(BorderStyle.THIN);// This is working
        styleMoney.setBorderBottom(BorderStyle.THIN);
        styleMoney.setBorderLeft(BorderStyle.THIN);
        styleMoney.setBorderRight(BorderStyle.THIN);

        //new style for all QB lines since they require wrap - 08/22/14
        styleMoneyWrap = m_Wrkbk.createCellStyle();
        styleMoneyWrap.setFont(font);
        styleMoneyWrap.setAlignment(HorizontalAlignment.RIGHT);
        styleMoneyWrap.setDataFormat((short) 8);
        styleMoneyWrap.setWrapText(true);

        styleMoneyWrap.setBorderTop(BorderStyle.THIN);// This is working
        styleMoneyWrap.setBorderBottom(BorderStyle.THIN);
        styleMoneyWrap.setBorderLeft(BorderStyle.THIN);
        styleMoneyWrap.setBorderRight(BorderStyle.THIN);

        m_CellStyles = new HSSFCellStyle[]{
                styleText,    // col 0 Message
                styleItemDesc,// col 1 Item Desc
                styleMoney,   // col 2 Reg Base Cost
                styleMoneyWrap,   // col 3 Promo Cost  -- changed to include line wrap - 08/22/14
                styleTextWrap,    // col 4 Stock Pack  -- changed to include line wrap - 08/22/14
                styleText,   // col 5 upc
                styleText,    // col 6 emery item #
                styleText,    // col 7 Item #
                styleText,    // col 8 Ord Qty
                styleText,    // col 9 Item #
                styleText,    // col 10 Ord Qty
        };

    }

    /**
     * Wraps the text to the no of lines specified(useful for long names e.g.item desc.)
     */
    private String[] wrapText(String text, int len) {
        //
        // return empty array for null text
        if (text == null)
            return new String[]{};

        //
        // return text if len is zero or less
        if (len <= 0)
            return new String[]{text};

        //
        // return text if less than length
        if (text.length() <= len)
            return new String[]{text};

        char[] chars = text.toCharArray();
        Vector<String> lines = new Vector<String>();
        StringBuffer line = new StringBuffer();
        StringBuffer word = new StringBuffer();

        for (int i = 0; i < chars.length; i++) {
            word.append(chars[i]);

            if (chars[i] == ' ') {
                if ((line.length() + word.length()) > len) {
                    lines.add(line.toString());
                    line.delete(0, line.length());
                }

                line.append(word);
                word.delete(0, word.length());
            }
        }

        //
        // handle any extra chars in current word
        if (word.length() > 0) {
            if ((line.length() + word.length()) > len) {
                lines.add(line.toString());
                line.delete(0, line.length());
            }
            line.append(word);
        }

        //
        // handle extra line
        if (line.length() > 0) {
            lines.add(line.toString());
        }

        String[] ret = new String[lines.size()];
        int c = 0;
        for (Enumeration<String> e = lines.elements(); e.hasMoreElements(); c++) {
            ret[c] = e.nextElement();
        }

        return ret;
    }

    /**
     * Removes forward or backward slashes from the vendor name.
     */
    private String backlashReplace(String vendName) {
        StringBuilder result = new StringBuilder();
        StringCharacterIterator itr = new StringCharacterIterator(vendName);
        char ch = itr.current();
        while (ch != CharacterIterator.DONE) {
            if (ch == '/' || ch == '\\') {
                result.append("-");
            } else {
                result.append(ch);
            }
            ch = itr.next();
        }
        return result.toString();
    }

    public static void main(String... args) {
        BasicConfigurator.configure();

        ShowPreprint rpt = new ShowPreprint();

        ArrayList<Param> params = new ArrayList<>();

        params.add(new Param("String", "2019 SPRING SHOW", "showdesc"));
        params.add(new Param("String", "124", "packetid"));
        params.add(new Param("String", "false", "extraentry"));
        params.add(new Param("String", "false", "hotbuy"));
        params.add(new Param("String", "false", "incredible"));
        params.add(new Param("String", "false", "nobrkcarton"));
        params.add(new Param("String", "false", "slamdunk"));
        //params.add(new Param("String", "", "customerid"));

        rpt.setParams(params);

        boolean created = rpt.createReport();

        System.out.println("Created? " + created);
    }

    private static Connection getConnectionPg() throws SQLException {
        java.util.Properties conProps = new java.util.Properties();
        conProps.put("user", "ejd");
        conProps.put("password", "boxer");

        return java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33/emery_jensen", conProps);
    }

}