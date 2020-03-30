/**
 * Title:			QuantityBuy.java
 * Description:   Quantity Buy Program Spreadsheet
 * Company:			Emery-Waterhouse
 *
 * @author prichter
 * @version 1.0 <p>
 * Create Date:	Mar 31, 2009
 * Last Update:   $Id: QuantityBuyWorksheet.java,v 1.1 2010/01/24 00:14:19 prichter Exp $
 * <p>
 * History:
 * $Log: QuantityBuyWorksheet.java,v $
 * Revision 1.1  2010/01/24 00:14:19  prichter
 * Initial add
 */
package com.emerywaterhouse.rpt.spreadsheet;

import com.emerywaterhouse.pricing.DiscountOrderLine;
import com.emerywaterhouse.pricing.DiscountWorksheet;
import com.emerywaterhouse.pricing.WorksheetItem;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.web.QuantityBuy;
import com.emerywaterhouse.websvc.Param;
import org.apache.log4j.BasicConfigurator;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class QuantityBuyWorksheet extends Report {

    // Parameters
    private String m_CustId;
    private String m_Title;
    private ArrayList<String> m_Warehouses;
    private ArrayList<Integer> m_Discounts;
    private ArrayList<String> m_ItemColumns;
    private ArrayList<String> m_LevelColumns;

    // Quantity buy related classes
    QuantityBuy m_QtyBuy;
    DiscountWorksheet m_Discount;

    // POI classes
    XSSFWorkbook m_Wrkbk = new XSSFWorkbook();
    XSSFSheet m_Sheet = m_Wrkbk.createSheet();
    private XSSFFont m_Font;
    private XSSFFont m_FontSubtitle;
    private XSSFFont m_FontTitle;
    private XSSFFont m_FontBold;
    private XSSFFont m_FontData;

    private XSSFCellStyle m_StyleTitle;    // size 12 Bold centered
    private XSSFCellStyle m_StyleSubtitle; // size 10 left-alligned
    private XSSFCellStyle m_StyleBoldText; //size 8 bold left-alligned
    private XSSFCellStyle m_StyleBoldNbr;  // size 8 bold right-alligned
    private XSSFCellStyle m_StyleBoldCtr;  // size 8 bold centered
    private XSSFCellStyle m_StyleText;    // Text right justified
    private XSSFCellStyle m_StyleDec;      // Style with 2 decimals
    private XSSFCellStyle m_StyleTextCtr;    // Text style centered
    private XSSFCellStyle m_StyleInt;      // Style with 0 decimals
    private XSSFCellStyle m_StylePct;      // Style with 0 decimals + %
    private XSSFCellStyle m_StyleLabel;    // Text labels, right justify, 8pt

    private int m_Row = 0;
    private int m_Col = 0;

    FileOutputStream m_OutFile;

    /**
     * Constructor
     */
    public QuantityBuyWorksheet() {
        super();

        m_Warehouses = new ArrayList<String>();
        m_Discounts = new ArrayList<Integer>();
        m_ItemColumns = new ArrayList<String>();
        m_LevelColumns = new ArrayList<String>();

        //
        // Create the default font for this workbook
        m_Font = m_Wrkbk.createFont();
        m_Font.setFontHeightInPoints((short) 8);
        m_Font.setFontName("Arial");

        //
        // Create a font that is normal size & bold
        m_FontBold = m_Wrkbk.createFont();
        m_FontBold.setFontHeightInPoints((short) 8);
        m_FontBold.setFontName("Arial");
        m_FontBold.setBold(true);

        //
        // Create a font that is normal size
        m_FontData = m_Wrkbk.createFont();
        m_FontData.setFontHeightInPoints((short) 8);
        m_FontData.setFontName("Arial");

        //
        // Create a font for sub titles
        m_FontSubtitle = m_Wrkbk.createFont();
        m_FontSubtitle.setFontHeightInPoints((short) 10);
        m_FontSubtitle.setFontName("Arial");
        m_FontSubtitle.setBold(true);

        //
        // Create a font for titles
        m_FontTitle = m_Wrkbk.createFont();
        m_FontTitle.setFontHeightInPoints((short) 12);
        m_FontTitle.setFontName("Arial");
        m_FontTitle.setBold(true);

        //
        // Setup the cell styles used in this report
        m_StyleBoldCtr = m_Wrkbk.createCellStyle();
        m_StyleBoldCtr.setFont(m_FontBold);
        m_StyleBoldCtr.setAlignment(HorizontalAlignment.RIGHT);
        m_StyleBoldCtr.setWrapText(true);

        m_StyleBoldText = m_Wrkbk.createCellStyle();
        m_StyleBoldText.setFont(m_FontBold);
        m_StyleBoldText.setAlignment(HorizontalAlignment.LEFT);
        m_StyleBoldText.setWrapText(true);

        m_StyleDec = m_Wrkbk.createCellStyle();
        m_StyleDec.setAlignment(HorizontalAlignment.RIGHT);
        m_StyleDec.setFont(m_FontData);
        m_StyleDec.setDataFormat((short) 4);

        m_StyleInt = m_Wrkbk.createCellStyle();
        m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
        m_StyleInt.setFont(m_FontData);
        m_StyleInt.setDataFormat((short) 3);

        m_StyleLabel = m_Wrkbk.createCellStyle();
        m_StyleLabel.setFont(m_Font);
        m_StyleLabel.setAlignment(HorizontalAlignment.RIGHT);

        m_StylePct = m_Wrkbk.createCellStyle();
        m_StylePct.setAlignment(HorizontalAlignment.RIGHT);
        m_StylePct.setFont(m_FontData);
        m_StylePct.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
        //m_StylePct.setDataFormat((short)4);


        m_StyleText = m_Wrkbk.createCellStyle();
        m_StyleText.setFont(m_FontData);
        m_StyleText.setAlignment(HorizontalAlignment.LEFT);
        m_StyleText.setWrapText(true);

        m_StyleTextCtr = m_Wrkbk.createCellStyle();
        m_StyleTextCtr.setFont(m_FontData);
        m_StyleTextCtr.setAlignment(HorizontalAlignment.RIGHT);
        m_StyleTextCtr.setWrapText(true);

        m_StyleSubtitle = m_Wrkbk.createCellStyle();
        m_StyleSubtitle.setFont(m_FontSubtitle);
        m_StyleSubtitle.setAlignment(HorizontalAlignment.LEFT);

        m_StyleTitle = m_Wrkbk.createCellStyle();
        m_StyleTitle.setFont(m_FontTitle);
        m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);
        m_StyleBoldNbr = m_Wrkbk.createCellStyle();
        m_StyleBoldNbr.setFont(m_FontBold);
        m_StyleBoldNbr.setAlignment(HorizontalAlignment.RIGHT);
        m_StyleBoldNbr.setWrapText(true);
    }

    /**
     * Clean up resources
     */
    public void close() {
        DbUtils.closeDbConn(m_OraConn, null, null);

        m_OraConn = null;

        m_Warehouses.clear();
        m_Discounts.clear();
        m_ItemColumns.clear();
        m_LevelColumns.clear();

        m_Warehouses = null;
        m_Discounts = null;
        m_ItemColumns = null;
        m_LevelColumns = null;
    }

    /**
     * Convenience method that adds a completely bordered cell with the specified alignment.
     *
     * @param rowNum - the row index.
     * @param colNum - the column index.
     * @return HSSFCell - the newly added cell, or a reference to the existing one.
     */
    private XSSFCell addBordCell(int rowNum, int colNum) 
    {
        return addCell(rowNum, colNum);
    }

    /**
     * Convenience method that adds a new String type cell with no borders and the specified alignment.
     *
     * @param rowNum int - the row index.
     * @param colNum short - the column index.
     * @param val String - the cell value.
     *
     * @return HSSFCell - the newly added String type cell, or a reference to the existing one.
     */
    private XSSFCell addCell(int rowNum, int colNum, String val, XSSFCellStyle style) 
    {
        XSSFCell cell = addCell(rowNum, colNum);

        cell.setCellType(CellType.STRING);
        cell.setCellValue(new HSSFRichTextString(val));
        cell.setCellStyle(style);

        return cell;
    }

    /**
     * Convenience method that adds a new numeric type cell with no borders and the specified alignment.
     *
     * @param rowNum - the row index.
     * @param colNum short - the column index.
     * @param val double - the cell value.
     * @param style XSSFCellStyle - the cell style and format
     *
     * @return HSSFCell - the newly added numeric type cell, or a reference to the existing one.
     */
    private XSSFCell addCell(int rowNum, int colNum, double val, XSSFCellStyle style) 
    {
        XSSFCell cell = addCell(rowNum, colNum);

        cell.setCellType(CellType.NUMERIC);
        cell.setCellStyle(style);
        cell.setCellValue(val);

        return cell;
    }

    /**
     * Adds a new cell with the specified borders and horizontal alignment.
     *
     * @param rowNum - the row index.
     * @param colNum short - the column index.
     *
     * @return HSSFCell - the newly added cell, or a reference to the existing one.
     */
    private XSSFCell addCell(int rowNum, int colNum) 
    {
        XSSFRow row = addRow(rowNum);
        XSSFCell cell = row.getCell(colNum);

        if (cell == null)
            cell = row.createCell(colNum);

        row = null;

        return cell;
    }

    /**
     * Build the spreadsheet for the current discount
     *
     * @param discount DiscountWorksheet - the DiscountWorksheet representing the current discount
     */
    private void addDiscount(DiscountWorksheet discount) {
        ArrayList<DiscountOrderLine> lines;
        DiscountOrderLine line;
        String name = "";
        WorksheetItem item = null;

        addCell(m_Row, 0, discount.getDescription(), m_StyleTitle);

        createHeadings(discount);

        lines = discount.getItems();

        if (lines != null) {
            for (int i = 0; i < lines.size(); i++) {
                line = lines.get(i);
                item = line.getWorksheetLevel(0);

                // If the item is not sold in one of the selected
                // warehouses, skip it
                if (!warehouseOk(item))
                    continue;

                addRow(m_Row++);
                m_Col = 0;

                // Add the item# and description.  These are always included.
                addCell(m_Row, m_Col++, line.getItemId(), m_StyleText);
                addCell(m_Row, m_Col++, line.getItemDescr(), m_StyleText);

                // Loop through the optional item level fields and add them as needed.
                // The user has the ability to sort the column when the report
                // is requested, so print them in the order in which they are
                // arranged in the ArrayList.
                for (int col = 0; col < m_ItemColumns.size(); col++) {
                    name = m_ItemColumns.get(col);

                    if (name.equals("Base Sell")) {
                        addCell(m_Row, m_Col++, item.getSell(), m_StyleDec);
                        continue;
                    }

                    if (name.equals("Cube")) {
                        addCell(m_Row, m_Col++, item.getCube(), m_StyleDec);
                        continue;
                    }

                    if (name.equals("Customer Regular Margin")) {
                        addCell(m_Row, m_Col++, (item.getCustRetail() - item.getRegularSell()) / item.getCustRetail(), m_StylePct);
                        continue;
                    }

                    if (name.equals("Customer Retail")) {
                        addCell(m_Row, m_Col++, item.getCustRetail(), m_StyleDec);
                        continue;
                    }

                    if (name.equals("Customer Regular Cost")) {
                        addCell(m_Row, m_Col++, item.getRegularSell(), m_StyleDec);
                        continue;
                    }

                    if (name.equals("Discount Invoice Label")) {
                        addCell(m_Row, m_Col++, discount.getInvoiceLabel(), m_StyleTextCtr);
                        continue;
                    }

                    if (name.equals("Emery Cost")) {
                        addCell(m_Row, m_Col++, item.getCost(), m_StyleDec);
                        continue;
                    }

                    if (name.equals("NBC")) {
                        addCell(m_Row, m_Col++, item.getNbcDescr(), m_StyleTextCtr);
                        continue;
                    }

                    if (name.equals("Pallet Quantity")) {
                        addCell(m_Row, m_Col++, item.getPalletQty(), m_StyleInt);
                        continue;
                    }

                    if (name.equals("Program Begin")) {
                        addCell(m_Row, m_Col++, discount.getBeginDate().toString(), m_StyleTextCtr);
                        continue;
                    }

                    if (name.equals("Program End")) {
                        if (discount.getEndDate() != null)
                            addCell(m_Row, m_Col++, discount.getEndDate().toString(), m_StyleTextCtr);
                        else
                            m_Col++;

                        continue;
                    }

                    if (name.equals("Retail C")) {
                        addCell(m_Row, m_Col++, item.getRetail(), m_StyleDec);
                        continue;
                    }

                    if (name.equals("Retail Pack")) {
                        addCell(m_Row, m_Col++, item.getRetailPack(), m_StyleInt);
                        continue;
                    }

                    if (name.equals("Retail Unit")) {
                        addCell(m_Row, m_Col++, item.getRetailUnit(), m_StyleTextCtr);
                        continue;
                    }

                    if (name.equals("Ship Unit")) {
                        addCell(m_Row, m_Col++, item.getShipUnit(), m_StyleTextCtr);
                        continue;
                    }

                    if (name.equals("Stock Pack")) {
                        addCell(m_Row, m_Col++, item.getStockPack(), m_StyleInt);
                        continue;
                    }

                    if (name.equals("UPC")) {
                        addCell(m_Row, m_Col++, item.getUpc(), m_StyleText);
                        continue;
                    }

                    if (name.equals("Vendor Name")) {
                        addCell(m_Row, m_Col++, item.getVendorName(), m_StyleText);
                        continue;
                    }

                    if (name.equals("Warehouses")) {
                        addCell(m_Row, m_Col++, getWarehouses(item), m_StyleTextCtr);
                        continue;
                    }

                    if (name.equals("Weight")) {
                        addCell(m_Row, m_Col++, item.getWeight(), m_StyleDec);
                        continue;
                    }
                }

                //
                // Loop through the levels and add level specific columns
                for (int j = 0; j < discount.getDiscountLevelCount(); j++) {
                    item = line.getWorksheetLevel(j);

                    if (item != null) {
                        for (int col = 0; col < m_LevelColumns.size(); col++) {
                            name = m_LevelColumns.get(col);

                            if (name.equals("Adders Apply")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().addersApply() ? "Y" : "N", m_StyleTextCtr);
                                continue;
                            }

                            if (name.equals("Break Amount")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().getBreakAmount(), m_StyleDec);
                                continue;
                            }

                            if (name.equals("Break Type")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().getBreakType(), m_StyleTextCtr);
                                continue;
                            }

                            if (name.equals("Cost Difference")) {
                                addCell(m_Row, m_Col++, item.getRegularSell() - item.getSell(), m_StyleDec);
                                continue;
                            }

                            if (name.equals("Discount Amount")) {
                                addCell(m_Row, m_Col++, item.getDiscountAmount(), m_StyleDec);
                                continue;
                            }

                            if (name.equals("Discount Method")) {
                                addCell(m_Row, m_Col++, item.getDiscountMethod(), m_StyleText);
                                continue;
                            }

                            if (name.equals("Discounted Customer Sell")) {
                                item.getOrderLine().setSellPrice(item.calcNewPrice());
                                addCell(m_Row, m_Col++, item.getOrderLine().getSellPrice(), m_StyleDec);
                                continue;
                            }

                            if (name.equals("Discounted Customer Margin")) {
                                addCell(m_Row, m_Col++, (item.getCustRetail() - item.getOrderLine().getSellPrice()) / item.getCustRetail(), m_StylePct);
                                continue;
                            }

                            if (name.equals("Emery Margin")) {
                                addCell(m_Row, m_Col++, (item.getOrderLine().getSellPrice() - item.getCost()) / item.getOrderLine().getSellPrice(), m_StylePct);
                                continue;
                            }

                            if (name.equals("Functional Discounts Apply")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().functionalDiscountsApply() ? "Y" : "N", m_StyleTextCtr);
                                continue;
                            }

                            if (name.equals("Minimum Quantity")) {
                                addCell(m_Row, m_Col++, item.getMinQuantity(), m_StyleDec);
                                continue;
                            }

                            if (name.equals("Mixed Items")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().canMixItems() ? "Y" : "N", m_StyleTextCtr);
                                continue;
                            }

                            if (name.equals("Packet Number")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().getPacketId(), m_StyleText);
                                continue;
                            }

                            if (name.equals("Promo Number")) {
                                addCell(m_Row, m_Col++, item.getDiscountLevel().getPromoId(), m_StyleText);
                                continue;
                            }

                            if (name.equals("Promo Payment Terms")) {
                                if (item.getDiscountLevel().getTermId() > 0)
                                    addCell(m_Row, m_Col, item.getDiscountLevel().getTermName(), m_StyleText);

                                m_Col++;
                                continue;
                            }

                            if (name.equals("Sell Multiple")) {
                                addCell(m_Row, m_Col++, item.getSellMultiple(), m_StyleInt);
                                continue;
                            }
                        }
                    }
                }
            }
        }

        m_Row++;
        m_Row++;
    }

    /**
     * Convenience method that adds a merged region with the given value and border style.  This method assumes
     * that the region will only have 1 row, and 1 or more cells.  Also assumes a string type value.
     *
     * @param rowNum - start from this row.
     * @param colFrom short - start from this column.
     * @param colTo short - to this column.
     * @param value String - region cell value.
     * @param merge boolean - if true then don't merge it.  This is here because of POI bug#16362. See bugzilla.
     *
     * Note:  The bug reference in the merge parameter is only an issue if the number of merged cells exceeds 1000.
     *        This isn't an issue in the
     */
    private void addRegion(int rowNum, int colFrom, int colTo, String value, boolean merge, XSSFCellStyle style) 
    {
        XSSFCell cell = null;

        addRow(rowNum);

        cell = addBordCell(rowNum, colFrom);
        cell.setCellType(CellType.STRING);
        cell.setCellValue(new HSSFRichTextString(value));
        cell.setCellStyle(style);

        if (merge) {
            CellRangeAddress region = new CellRangeAddress(rowNum, rowNum, colFrom, colTo);
            m_Sheet.addMergedRegion(region);
        }
    }

    /**
     * Adds a new row or returns the existing one.
     *
     * @param rowNum int - the row index.
     * @return HSSFRow - the row object added, or a reference to the existing one.
     */
    private XSSFRow addRow(int rowNum) 
    {
        XSSFRow row = m_Sheet.getRow(rowNum);

        if (row == null)
            row = m_Sheet.createRow(rowNum);

        return row;
    }

    /**
     * Creates the output file
     * @return boolean - true if successful
     * @throws FileNotFoundException
     * @throws SQLException
     */
    public boolean buildOutputFile() throws FileNotFoundException, SQLException {
        boolean built = false;
        StringBuffer fileName = new StringBuffer();
        String tmp = Long.toString(System.currentTimeMillis());

        fileName.append("QtyBuyWorksheet");
        fileName.append("-");
        fileName.append(tmp.substring(tmp.length() - 5, tmp.length()));
        fileName.append(".xls");
        m_FileNames.add(fileName.toString());
        m_OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
        fileName = null;

        m_QtyBuy = new QuantityBuy(m_OraConn, m_EdbConn);

        m_Sheet.getHeader().setLeft(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 10) + HeaderFooter.date());
        m_Sheet.getHeader().setCenter(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 12) + m_Title);
        m_Sheet.getHeader().setRight(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 12) + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        m_Row = 0;

        if (m_CustId != null && m_CustId.trim().length() == 6) {
            addCell(m_Row, 1, "Customer: " + m_CustId + " " + getCustName(m_CustId), m_StyleBoldText);
            m_Row++;
            m_Row++;
        }

        try {
            for (int i = 0; i < m_Discounts.size(); i++) {
                try {
                    m_Discount = m_QtyBuy.loadDiscountWorksheet(m_Discounts.get(i), m_CustId);

                    if (m_Discount != null) {
                        addDiscount(m_Discount);
                    }
                } catch (Exception e) {
                    log.error("Error in Quantity Buy Worksheet for discount: " +
                            m_Discounts.get(i) + " customer: " + m_CustId);
                    log.error("exception", e);
                } finally {
                    m_Discount = null;
                }
            }

            m_Wrkbk.write(m_OutFile);
            m_OutFile.close();

            built = true;
        } catch (Exception e) {
            log.error("exception", e);
            built = false;
        } finally {
            m_Wrkbk = null;
            m_Discount = null;
            m_OutFile = null;
        }

        return built;
    }

    /**
     * Builds the report and column headings
     *
     * @param discount - the current DiscountWorksheet
     */
    public void createHeadings(DiscountWorksheet discount) {
        String name;

        m_Col = 0;
        m_Row++;

        m_Sheet.setColumnWidth(m_Col, 2000);
        addCell(m_Row, m_Col++, "Item#", m_StyleBoldText);

        m_Sheet.setColumnWidth(m_Col, 12000);
        addCell(m_Row, m_Col++, "Description", m_StyleBoldText);

        for (int col = 0; col < m_ItemColumns.size(); col++) {
            name = m_ItemColumns.get(col);

            if (name.equals("Base Sell")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Base Sell", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Cube")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Unit Cube", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Customer Regular Margin")) {
                m_Sheet.setColumnWidth(m_Col, 2300);
                addCell(m_Row, m_Col++, "Customer Reg Mgn", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Customer Retail")) {
                m_Sheet.setColumnWidth(m_Col, 2300);
                addCell(m_Row, m_Col++, "Customer Retail", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Customer Regular Cost")) {
                m_Sheet.setColumnWidth(m_Col, 2300);
                addCell(m_Row, m_Col++, "Customer Reg Cost", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Discount Invoice Label")) {
                m_Sheet.setColumnWidth(m_Col, 2500);
                addCell(m_Row, m_Col++, "Program Name", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Emery Cost")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Emery Cost", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("NBC")) {
                m_Sheet.setColumnWidth(m_Col, 1000);
                addCell(m_Row, m_Col++, "NBC", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Pallet Quantity")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Pallet Qty", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Program Begin")) {
                m_Sheet.setColumnWidth(m_Col, 2500);
                addCell(m_Row, m_Col++, "Program Begin", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Program End")) {
                m_Sheet.setColumnWidth(m_Col, 2500);
                addCell(m_Row, m_Col++, "Program End", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Retail C")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "RetailC", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Retail Pack")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Retail Pack", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("Retail Unit")) {
                m_Sheet.setColumnWidth(m_Col, 2000);
                addCell(m_Row, m_Col++, "Retail UOM", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Ship Unit")) {
                m_Sheet.setColumnWidth(m_Col, 2000);
                addCell(m_Row, m_Col++, "Ship UOM", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Stock Pack")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Stock Pack", m_StyleBoldNbr);
                continue;
            }

            if (name.equals("UPC")) {
                m_Sheet.setColumnWidth(m_Col, 3000);
                addCell(m_Row, m_Col++, "UPC", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Vendor Name")) {
                m_Sheet.setColumnWidth(m_Col, 8000);
                addCell(m_Row, m_Col++, "Manufacturer", m_StyleBoldText);
                continue;
            }

            if (name.equals("Warehouses")) {
                m_Sheet.setColumnWidth(m_Col, 3000);
                addCell(m_Row, m_Col++, "Warehouses", m_StyleBoldCtr);
                continue;
            }

            if (name.equals("Weight")) {
                m_Sheet.setColumnWidth(m_Col, 1500);
                addCell(m_Row, m_Col++, "Unit Wgt", m_StyleBoldNbr);
                continue;
            }
        }

        //
        // Loop through the levels and add headings for each level specific column
        for (int j = 0; j < discount.getDiscountLevelCount(); j++) {
            //If there is more than 1 column per level, create a merged
            //region above the level column headings for the level
            //description.  If there is only 1, place the level heading
            //above the column heading.  If there are none, don't use
            //the heading at all.
            if (m_LevelColumns.size() > 1)
                addRegion(m_Row - 1, m_Col, (m_Col + m_LevelColumns.size() - 1), discount.getDiscountLevel(j).getLevelDescr(), true, m_StyleBoldCtr);
            else if (m_LevelColumns.size() == 1)
                addCell(m_Row - 2, m_Col, discount.getDiscountLevel(j).getLevelDescr(), m_StyleBoldCtr);

            for (int col = 0; col < m_LevelColumns.size(); col++) {
                name = m_LevelColumns.get(col);

                if (name.equals("Adders Apply")) {
                    m_Sheet.setColumnWidth(m_Col, 1700);
                    addCell(m_Row, m_Col++, "Adders Apply", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Break Amount")) {
                    m_Sheet.setColumnWidth(m_Col, 2000);
                    addCell(m_Row, m_Col++, "Break Point", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Break Type")) {
                    m_Sheet.setColumnWidth(m_Col, 3000);
                    addCell(m_Row, m_Col++, "Break Type", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Cost Difference")) {
                    m_Sheet.setColumnWidth(m_Col, 2000);
                    addCell(m_Row, m_Col++, "Cost Diff", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Discount Amount")) {
                    m_Sheet.setColumnWidth(m_Col, 2000);
                    addCell(m_Row, m_Col++, "Discount Value", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Discount Method")) {
                    m_Sheet.setColumnWidth(m_Col, 2500);
                    addCell(m_Row, m_Col++, "Discount Method", m_StyleBoldText);
                    continue;
                }

                if (name.equals("Discounted Customer Sell")) {
                    m_Sheet.setColumnWidth(m_Col, 2500);
                    addCell(m_Row, m_Col++, "Discounted Sell", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Discounted Customer Margin")) {
                    m_Sheet.setColumnWidth(m_Col, 2500);
                    addCell(m_Row, m_Col++, " Discounted Margin", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Emery Margin")) {
                    m_Sheet.setColumnWidth(m_Col, 2000);
                    addCell(m_Row, m_Col++, "Emery Margin", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Functional Discounts Apply")) {
                    m_Sheet.setColumnWidth(m_Col, 1500);
                    addCell(m_Row, m_Col++, "Func Disc Apply", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Minimum Quantity")) {
                    m_Sheet.setColumnWidth(m_Col, 1800);
                    addCell(m_Row, m_Col++, "Min Qty", m_StyleBoldNbr);
                    continue;
                }

                if (name.equals("Mixed Items")) {
                    m_Sheet.setColumnWidth(m_Col, 1700);
                    addCell(m_Row, m_Col++, "Mixed SKUs", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Packet Number")) {
                    m_Sheet.setColumnWidth(m_Col, 1500);
                    addCell(m_Row, m_Col++, "Pckt", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Promo Number")) {
                    m_Sheet.setColumnWidth(m_Col, 1600);
                    addCell(m_Row, m_Col++, "Promo", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Promo Payment Terms")) {
                    m_Sheet.setColumnWidth(m_Col, 4000);
                    addCell(m_Row, m_Col++, "Promotional Payment Terms", m_StyleBoldCtr);
                    continue;
                }

                if (name.equals("Sell Multiple")) {
                    m_Sheet.setColumnWidth(m_Col, 2000);
                    addCell(m_Row, m_Col++, "Sell Multiple", m_StyleBoldNbr);
                    continue;
                }
            }
        }
    }

    /* (non-Javadoc)
     * @see com.emerywaterhouse.rpt.server.Report#createReport()
     */
    @Override
    public boolean createReport() {
        boolean created = false;
        m_Status = RptServer.RUNNING;

        try {
            m_OraConn = m_RptProc.getOraConn();
            m_EdbConn = m_RptProc.getEdbConn();
            
            created = buildOutputFile();
        } catch (Exception ex) {
            log.fatal("exception:", ex);
        } finally {
            close();

            if (m_Status == RptServer.RUNNING)
                m_Status = RptServer.STOPPED;
        }

        return created;
    }

    /**
     * Returns the name of the customer the spreadsheet it run for
     *
     * @param custId String - the customer id
     * @return String - the custome name
     * @throws SQLException
     */
    private String getCustName(String custId) throws SQLException {
        String name = "Invalid Customer#";
        Statement stmt = null;
        ResultSet rs = null;

        try {
            stmt = m_OraConn.createStatement();
            rs = stmt.executeQuery("select name from customer where customer_id = '" + custId + "'");

            if (rs.next())
                name = rs.getString("name");
        } finally {
            DbUtils.closeDbConn(null, stmt, rs);
            rs = null;
            stmt = null;
        }

        return name;
    }

    /**
     * Returns the list of warehouses this item is stocked in
     *
     * @param item WorksheetItem
     * @return String - a list of warehouse names separated by spaces
     */
    private String getWarehouses(WorksheetItem item) {
        String whsList = "";

        for (int i = 0; i < item.getWarehouseCount(); i++) {
            if (whsList.length() == 0)
                whsList = item.getWarehouse(i);
            else
                whsList = whsList + " " + item.getWarehouse(i);
        }

        return whsList;
    }

    /**
     * Sets the parameters of this report.
     * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
     */
    public void setParams(ArrayList<Param> params) {
        String name;

        for (int i = 0; i < params.size(); i++) {
            name = params.get(i).name;

            if (name.equals("custid")) {
                m_CustId = params.get(i).value.trim();
                continue;
            }

            if (name.equals("title")) {
                m_Title = params.get(i).value.trim();
                continue;
            }

            if (name.equals("warehouse")) {
                m_Warehouses.add(params.get(i).value);
                continue;
            }

            if (name.equals("discount")) {
                m_Discounts.add(new Integer(params.get(i).value));
                continue;
            }

            if (name.equals("item column")) {
                m_ItemColumns.add(params.get(i).value);
                continue;
            }

            if (name.equals("level column")) {
                m_LevelColumns.add(params.get(i).value);
                continue;
            }
        }
    }

    /**
     * Checks the warehouses the item is stocked in against the warehouses
     * requested in the report.  Returns true if the item is stocked
     * in one of the selected warehouses.
     *
     * @param item String - the item id
     * @return boolean - true if the item is stocked in one of the selected warehouses
     */
    private boolean warehouseOk(WorksheetItem item) {
        boolean ok = false;

        for (int i = 0; i < m_Warehouses.size(); i++) {
            for (int j = 0; j < item.getWarehouseCount(); j++) {
                if (m_Warehouses.get(i).equals(item.getWarehouse(j))) {
                    ok = true;
                    break;
                }
            }
        }

        return ok;
    }

    public static void main(String... args) {
        BasicConfigurator.configure();

        ArrayList<Param> params = new ArrayList<>();

        //cust id
        Param p = new Param("String", "065994", "custid");
        params.add(p);

        //title
        p = new Param("String", "Quantity Buy Report", "title");
        params.add(p);

        //warehouses
        p = new Param("String", "PORTLAND", "warehouse");
        params.add(p);
        p = new Param("String", "LOXLEY", "warehouse");
        params.add(p);
        p = new Param("String", "LITTLE ROCK", "warehouse");
        params.add(p);
        p = new Param("String", "PRESCOTT VALLEY", "warehouse");
        params.add(p);
        p = new Param("String", "COLORADO SPRINGS", "warehouse");
        params.add(p);
        p = new Param("String", "TAMPA", "warehouse");
        params.add(p);
        p = new Param("String", "GAINESVILLE", "warehouse");
        params.add(p);
        p = new Param("String", "PRINCETON", "warehouse");
        params.add(p);
        p = new Param("String", "WILTON", "warehouse");
        params.add(p);
        p = new Param("String", "WEST JEFFERSON", "warehouse");
        params.add(p);
        p = new Param("String", "WILMER", "warehouse");
        params.add(p);
        p = new Param("String", "PRINCE GEORGE", "warehouse");
        params.add(p);
        p = new Param("String", "MOXEE", "warehouse");
        params.add(p);
        p = new Param("String", "LA CROSSE", "warehouse");
        params.add(p);
        p = new Param("String", "HILLMAN", "warehouse");
        params.add(p);
        p = new Param("String", "MIDWEST", "warehouse");
        params.add(p);
        p = new Param("String", "BRASSCRAFT", "warehouse");
        params.add(p);

        //Discounts
        p = new Param("String", "1574", "discount");
        params.add(p);
        p = new Param("String", "1121", "discount");
        params.add(p);
        p = new Param("String", "1938", "discount");
        params.add(p);
        p = new Param("String", "594", "discount");
        params.add(p);
        p = new Param("String", "1032", "discount");
        params.add(p);
        p = new Param("String", "1478", "discount");
        params.add(p);
        p = new Param("String", "1863", "discount");
        params.add(p);

        //item columns
        p = new Param("String", "Customer Retail", "item column");
        params.add(p);
        p = new Param("String", "UPC", "item column");
        params.add(p);

        //level columns
        p = new Param("String", "Sell Multiple", "level column");
        params.add(p);

        QuantityBuyWorksheet test = new QuantityBuyWorksheet();

        test.setParams(params);

        System.out.println(test.createReport());
    }

    /**
     * Attempts to create and return a connection to Grok
     */
    public static Connection getConnectionGrok() throws SQLException {
        Connection conn;
        java.util.Properties connProps = new java.util.Properties();
        connProps.put("user", "eis_emery");
        connProps.put("password", "boxer");
        conn = java.sql.DriverManager.getConnection(
                "jdbc:oracle:thin:@10.128.0.9:1521:GROK", connProps);

        return conn;
    }

}

