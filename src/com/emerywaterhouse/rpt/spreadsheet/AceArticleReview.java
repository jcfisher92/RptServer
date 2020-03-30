package com.emerywaterhouse.rpt.spreadsheet;

import java.sql.*;
import java.util.*;
import java.util.Date;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddressList;

import com.emerywaterhouse.rpt.server.*;
import com.emerywaterhouse.websvc.Param;

public class AceArticleReview extends Report {
   
   private static final String 
      BASE_QUERY = "select item.vendor_id, ace_item_xref.ace_sku, vendor.name vendor_name, " +
                   "  item.item_id, item.description, ace_item_xref.dept_num ace_dept, upc_code, " +
                   "  buyer_cd ace_buyer_code, stock_pack, retail_pack, broken_case_id, item.weight, " +
                   "  ace_item_dup.old_item_id emery_match, item.disp_id, item_disp.disposition, isr.reason, " +
                   "  null \"APPROVE/IGNORE\" " +
                   "from ace_item_xref " +
                   "join item on item.item_id = ace_item_xref.item_id " +
                   "join vendor on vendor.vendor_id = item.vendor_id " +
                   "join item_disp on item.disp_id = item_disp.disp_id " +
                   "join ace_item_rsc using (ace_xref_id) " +
                   "left outer join ace_item_dup on ace_item_dup.new_item_id = item.item_id " +
                   "left outer join item_upc on item_upc.item_id = item.item_id and primary_upc = 1 " +
                   "left outer join ace_item_status_reason isr on isr.ace_isr_id = ace_item_xref.ace_isr_id " +
                   "where item_type_id = 8 and ace_rsc_id = 11 and ace_item_xref.buyer_cd in (", // item_type_id 8 = 'ACE'; ace_rsc_id 11 = wilton, ny
      BASE_QUERY_DISP = " and disposition in (",
      BASE_QUERY_ORDER_BY = "order by item.disp_id DESC, vendor.name, ace_item_xref.dept_num ",
      PRICING_QUERY = "select item_id, reg_cost buy, ace_retailer_cost, ip.sell, retail_c retail " +
                      "from ace_item_price ip " +
                      "join ace_item_xref using (item_id) " +
                      "join item using (item_id) " +
                      "where ace_item_xref.buyer_cd in ( ",
      ACE_FREIGHT_ADDER_QUERY = "select adder_rate from ace_frt_adder where sell = ? and weight = ?",
      ORGILL_PRICING_QUERY = "select upc, orgill_price, qty_round, decode(broken_case_id, 1, 0, 1) nbc " +
                             "from item_orgill_price " +
                             "join item_upc on primary_upc = 1 and item_orgill_price.upc = item_upc.upc_code " +
                             "join ace_item_xref using (item_id) " +
                             "join item using (item_id) " +
                             "where buyer_cd in( ", 
      DISPOSITION_REASON_QUERY = "select reason from ace_item_status_reason ";
   
   private static final double 
      ORGILL_ADDER_RATE = 0.037,
      ORGILL_BROKEN_PACK_CHARGE_RATE = 0.04;

   private static final String
      DECIMAL_FORMAT = "#,##0.000;[Red]-(#,##0.000)",
      MONEY_FORMAT = "$#,##0.00;[Red]-($#,##0.00)",
      PERCENT_FORMAT = "0.000%;[Red]-0.000%";
   
   private static final short FONT_SIZE = 11;
   private static final String FONT_NAME = "Calibri";
   private static final float HEADING_ROW_HEIGHT_IN_POINTS = 53.5f;
   
   // if rptserver classloader gets fixed all this mess can be cleaned up
   private static final int 
      HEADING = 0,
      STYLE = 1;

   private static final int
      STYLE_TEXT = 0,
      STYLE_ID = 1,
      STYLE_INTEGER = 2,
      STYLE_FLOAT = 3,
      STYLE_MONEY = 4,
      STYLE_HEADING = 5,
      STYLE_PERCENT = 6;
   
   private static final Object[][] 
      m_BaseColumns = {
         {"VENDOR_ID", STYLE_ID},
         {"VENDOR_NAME", STYLE_TEXT},
         {"ITEM_ID", STYLE_ID},
         {"DESCRIPTION", STYLE_TEXT},
         {"EMERY_MATCH", STYLE_ID},
         {"UPC_CODE", STYLE_ID},
         {"ACE_SKU", STYLE_ID},
         {"ACE_DEPT", STYLE_ID},
         {"ACE_BUYER_CODE", STYLE_ID},
         {"STOCK_PACK", STYLE_INTEGER},
         {"RETAIL_PACK", STYLE_INTEGER},
         {"BROKEN_CASE_ID", STYLE_ID},
         {"WEIGHT", STYLE_FLOAT}
      },
      m_PricingColumns = {
         {"BUY", STYLE_MONEY},
         {"ACE_RETAILER_COST", STYLE_MONEY},
         {"SELL", STYLE_MONEY},
         {"ADDER_RATE", STYLE_FLOAT},
         {"FREIGHT", STYLE_FLOAT},
         {"NBC_FRT", STYLE_FLOAT},
         {"UPS_FRT", STYLE_FLOAT},
         {"FRT_DIFF", STYLE_FLOAT},
         {"NEW_SELL", STYLE_MONEY},
         {"RETAIL", STYLE_MONEY},
         {"RETAIL_MGN", STYLE_FLOAT}
      },
      m_OrgillPricingColumns = {
         {"ORG_ADDERS", STYLE_MONEY},
         {"E-O_DIFF", STYLE_FLOAT},
         {"E-O_PERCENT", STYLE_PERCENT}
      },
      m_AceDispositionColumns = {
         {"DISPOSITION", STYLE_TEXT},
         {"DISP_ID", STYLE_ID}
      },
      m_ReviewColumns = {
         {"REASON", STYLE_TEXT},
         {"APPROVE/IGNORE", STYLE_TEXT}
      };
   
   private static final Object[][][] m_ColumnGroups = new Object[][][]{m_BaseColumns, m_PricingColumns, 
         m_OrgillPricingColumns, m_AceDispositionColumns, m_ReviewColumns};
   
   private int[] 
      m_AceFreightSellPrices,
      m_AceFreightWeights;
   
   private Workbook m_Workbook;
   private Sheet m_Sheet;
   private List<CellStyle> m_Styles;
   private ResultSet 
      m_ItemBaseData,
      m_ItemPricingData,
      m_OrgillItemPricingData;
   private PreparedStatement 
      m_SelectItemBaseData,
      m_SelectItemPricingData,
      m_SelectOrgillPrices,
      m_SelectDispositionReason,
      m_SelectAceFreightAdderRate;
   
   private Map<String, Integer> 
      m_ItemPricingMap,
      m_OrgillItemPricingMap;
   private List<String> 
      m_BuyerCodeList,
      m_DispositionList;

   private Row m_CurrentRow;
   
   
   public AceArticleReview() {
      m_Styles = new ArrayList<>();
      m_Workbook = new XSSFWorkbook();
      m_ItemPricingMap = new HashMap<>();
      m_OrgillItemPricingMap = new HashMap<>();
   }

   @Override
   public boolean createReport() {
      boolean created = false;
      try {
         m_Status = RptServer.RUNNING;
         loadData();
         buildNewSheet();
         writeDataToWorkbook();
         outputWorkbook();
         created = true;
      } catch (Exception e) {
         m_ErrMsg.append("Your 'Ace Article Review' had the following errors: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n");
         m_ErrMsg.append(e.getMessage());
         log.error("[AceArticleReview]", e);
      } finally {
         m_Status = RptServer.STOPPED;
      }
      return created;
   }
   
   private void loadData() throws Exception {      
      prepareStatements();
      buildResultSets();
      loadPricingData();
      loadOrgillPricingData();
   }
   
   private void buildNewSheet() throws Exception {
      m_Sheet = m_Workbook.createSheet();
      buildCellStyles();
      m_Sheet.setZoom(3);
      buildColumnHeadings();
   }

   private void writeDataToWorkbook() throws Exception {
      while (m_ItemBaseData.next()) {
         writeRowToWorkbook(m_ItemBaseData.getRow());
      }
      
      int columnCount = m_Sheet.getRow(0).getPhysicalNumberOfCells();
      for (int i = 0; i < columnCount; i++)
         m_Sheet.autoSizeColumn(i);
      
      buildDropdownLists();
   }
   
   private void outputWorkbook() throws Exception {
      String fileName = new SimpleDateFormat("'Ace_Article_Review'yyMMddHHmmssS'.xlsx'").format(new Date());
      m_FileNames.add(fileName);
      FileOutputStream outputFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      m_Workbook.write(outputFile);
   }
   

   private void prepareStatements() throws Exception {
      String baseQuery = buildWhereClause(BASE_QUERY, m_BuyerCodeList) + buildWhereClause(BASE_QUERY_DISP, m_DispositionList) + BASE_QUERY_ORDER_BY;
      m_SelectItemBaseData = m_EdbConn.prepareStatement(baseQuery, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
      
      String pricingQuery = buildWhereClause(PRICING_QUERY, m_BuyerCodeList);
      m_SelectItemPricingData = m_EdbConn.prepareStatement(pricingQuery, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
      
      String orgillPricingQuery = buildWhereClause(ORGILL_PRICING_QUERY, m_BuyerCodeList);
      m_SelectOrgillPrices = m_EdbConn.prepareStatement(orgillPricingQuery, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
      
      m_SelectDispositionReason = m_EdbConn.prepareStatement(DISPOSITION_REASON_QUERY, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
      
      m_SelectAceFreightAdderRate = m_EdbConn.prepareStatement(ACE_FREIGHT_ADDER_QUERY);
   }
   
   private void buildResultSets() throws Exception {
      m_ItemBaseData = fillInWhereClauseParameters(m_SelectItemBaseData).executeQuery();
      m_ItemPricingData = fillInBuyerCodes(m_SelectItemPricingData).executeQuery();
      m_OrgillItemPricingData = fillInBuyerCodes(m_SelectOrgillPrices).executeQuery(); 
   }

   private void loadPricingData() throws Exception {
      while (m_ItemPricingData.next())
         m_ItemPricingMap.put(m_ItemPricingData.getString("ITEM_ID"), m_ItemPricingData.getRow());
      
      m_AceFreightSellPrices = getDistinctAceFreightSellPrices();
      m_AceFreightWeights = getDistinctAceFreightWeights();
   }

   private void loadOrgillPricingData() throws SQLException {
      while (m_OrgillItemPricingData.next())
         m_OrgillItemPricingMap.put(m_OrgillItemPricingData.getString("UPC"), m_OrgillItemPricingData.getRow());
   }

   private void buildColumnHeadings() throws Exception {
      m_Sheet.createRow(0).setHeightInPoints(HEADING_ROW_HEIGHT_IN_POINTS);
      m_Sheet.createFreezePane(0, 1);
         
      for (Object[][] columnGroup : m_ColumnGroups) {
         buildHeadingGroup(columnGroup);
      }
   }

   private String buildWhereClause(String query, List<String> parameters) {
      StringBuilder whereClauseBuilder = new StringBuilder(query);
      int lastIndex = parameters.size() - 1;
      
      for (int i = 0; i <= lastIndex; i++)
         whereClauseBuilder.append(i == lastIndex ? "?) " : "?,");

      return whereClauseBuilder.toString();
   }

   private PreparedStatement fillInWhereClauseParameters(PreparedStatement statement) throws Exception {
      int parameterIndex = 1;
      for (String code : m_BuyerCodeList) {
         statement.setString(parameterIndex++, code);
      }
      
      for (String disposition : m_DispositionList) {
         statement.setString(parameterIndex++, disposition);
      }
      return statement;
   }
   
   private PreparedStatement fillInBuyerCodes(PreparedStatement statement) throws Exception {
      int parameterIndex = 1;
      for (String code : m_BuyerCodeList) {
         statement.setString(parameterIndex++, code);
      }
      return statement;
   }

   private void writeRowToWorkbook(int rowIndex) throws Exception {
      m_CurrentRow = m_Sheet.createRow(rowIndex);
      
      String 
         itemUPC = m_ItemBaseData.getString("UPC_CODE"),
         itemID = m_ItemBaseData.getString("ITEM_ID");
      
      writeBaseColumnEntries();
      
      if (m_ItemPricingMap.containsKey(itemID)) {
         writePricingColumnEntries(itemID);
         
         if (m_OrgillItemPricingMap.containsKey(itemUPC))
            writeOrgillPricingColumnEntries(itemUPC);
      }
      
      writeAceDispositionColumnEntries();
      
      if (m_ItemBaseData.getString("REASON") != null)
         writeDispositionReasonColumnEntries();
   }

   private void buildDropdownLists() throws Exception {
      int columnIndex = getFirstColumnIndex(m_ReviewColumns);
      
      buildDropdownList(columnIndex, loadDispositionReasons());
      buildDropdownList(++columnIndex, new String[]{"Approve", "Ignore"});
   }
   
   private void buildDropdownList(int columnIndex, String[] listContents) {
      DataValidationHelper validationHelper = m_Sheet.getDataValidationHelper();
      CellRangeAddressList addressList = new CellRangeAddressList(1, m_Sheet.getLastRowNum(), columnIndex, columnIndex);
      DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(listContents);
      DataValidation validation = validationHelper.createValidation(constraint, addressList);
      
      m_Sheet.addValidationData(validation);
   }
   
   private void buildHeadingGroup(Object[][] group) throws Exception {
      int columnIndex = getNextUnwrittenCellIndex(m_Sheet.getRow(0));
      for (Object[] column : group) {
         String heading = (String)column[HEADING];
         Cell cell = m_Sheet.getRow(0).createCell(columnIndex++);
         
         cell.setCellStyle(m_Styles.get(STYLE_HEADING));
         writeCellValue(heading, cell);
      }
   }
   
   private void writeBaseColumnEntries() throws Exception {
      writeColumnEntries(m_BaseColumns, m_ItemBaseData);
   }
   
   private void writeColumnEntries(Object[][] columnGroup, ResultSet data) throws SQLException {
      int columnIndex = getFirstColumnIndex(columnGroup);
      
      for (Object[] column : columnGroup) {
         Object value = data.getObject((String)column[HEADING]);
         
         final Cell cell = m_CurrentRow.createCell(columnIndex++);
         if (value != null) {
            CellStyle style = m_Styles.get((int)column[STYLE]);
            cell.setCellStyle(style);
            
            writeCellValue(value, cell);
         }
      }
   }
   
   private void writeColumnEntries(Object[][] columnGroup, Map<String, Number> data) {
      int columnIndex = getFirstColumnIndex(columnGroup);
      
      for (Object[] column : columnGroup) {
         Number value = data.get((String)column[HEADING]);
         Cell cell = m_CurrentRow.createCell(columnIndex++);
         
         cell.setCellStyle(m_Styles.get((int)column[STYLE]));
         writeCellValue(value, cell);
      }
   }
   
   private void writePricingColumnEntries(String currentItem) throws Exception {
      final int pricingDataRowIndex = (m_ItemPricingMap.get(currentItem));
      m_ItemPricingData.absolute(pricingDataRowIndex);
      
      writeColumnEntries(m_PricingColumns, calculatePricingData());
   }

   private void writeOrgillPricingColumnEntries(String itemUPC) throws Exception {
      final int pricingDataRowIndex = (m_OrgillItemPricingMap.get(itemUPC));
      m_OrgillItemPricingData.absolute(pricingDataRowIndex);
      
      writeColumnEntries(m_OrgillPricingColumns, calculateOrgillPricingData());
   }
   
   private Map<String, Number> calculatePricingData() throws Exception {
      double
         adjustmentAdderRate = 0.95,
         costToEmery = m_ItemPricingData.getDouble("BUY"),
         adjustedCostToEmery = costToEmery / adjustmentAdderRate,
         aceSell = m_ItemPricingData.getDouble("SELL"),
         
         itemWeight = m_ItemBaseData.getDouble("WEIGHT"),
         aceFreightAdderRate = calculateAceFreightAdderRate(aceSell, itemWeight),
         freightCost = adjustedCostToEmery * aceFreightAdderRate,
         nonBreakableCaseFreightCost = freightCost * m_ItemBaseData.getDouble("STOCK_PACK"),
         
         lineCoefficient = 0.1955,
         lineConstant = 7.9581,
         upsFreightCost = lineCoefficient*itemWeight + lineConstant,
         freightCostDifference = nonBreakableCaseFreightCost - upsFreightCost,
         
         adjustedCostToCustomer = adjustedCostToEmery + freightCost,
         retailCost = m_ItemPricingData.getDouble("RETAIL"),
         retailMargin = (retailCost == 0) ? 0 : 1 - ((adjustedCostToCustomer/m_ItemBaseData.getDouble("RETAIL_PACK")) / retailCost);
      
      Map<String, Number> itemPricingData = new HashMap<>();
      itemPricingData.put("BUY", costToEmery);
      itemPricingData.put("ACE_RETAILER_COST", m_ItemPricingData.getDouble("ACE_RETAILER_COST"));
      itemPricingData.put("SELL", aceSell);
      itemPricingData.put("ADDER_RATE", aceFreightAdderRate);
      itemPricingData.put("FREIGHT", freightCost);
      itemPricingData.put("NBC_FRT", nonBreakableCaseFreightCost);
      itemPricingData.put("UPS_FRT", upsFreightCost);
      itemPricingData.put("FRT_DIFF", freightCostDifference);
      itemPricingData.put("NEW_SELL", adjustedCostToCustomer);
      itemPricingData.put("RETAIL", retailCost);
      itemPricingData.put("RETAIL_MGN", retailMargin);
      
      return itemPricingData;
   }
   
   private double calculateAceFreightAdderRate(double itemPrice, double itemWeight) throws Exception {
      int i = 0;
      int roundedSellPrice = m_AceFreightSellPrices[i++];
      while (i < m_AceFreightSellPrices.length && m_AceFreightSellPrices[i] <= itemPrice)
         roundedSellPrice = m_AceFreightSellPrices[i++];
      
      Arrays.sort(m_AceFreightWeights);
      i = 0;
      int roundedWeight = m_AceFreightWeights[i++];
      while (i < m_AceFreightWeights.length && m_AceFreightWeights[i] <= itemWeight)
         roundedWeight = m_AceFreightWeights[i++];
      
      m_SelectAceFreightAdderRate.setInt(1, roundedSellPrice);
      m_SelectAceFreightAdderRate.setInt(2, roundedWeight);
      ResultSet aceFreightAdderRate = m_SelectAceFreightAdderRate.executeQuery();
      aceFreightAdderRate.next();
      
      return aceFreightAdderRate.getDouble(1);
   }
  
   private int[] getDistinctAceFreightSellPrices() throws Exception {
      ResultSet sellPrices = m_OraConn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY)
            .executeQuery("select distinct sell from ace_frt_adder");
      
      int[] distinctSellPrices = new int[getResultSetRowCount(sellPrices)];
      for (int i=0; i<distinctSellPrices.length; i++) {
         sellPrices.next();
         distinctSellPrices[i] = sellPrices.getInt(1);
      }
      
      return distinctSellPrices;
   }
   
   private int[] getDistinctAceFreightWeights() throws Exception {
      ResultSet weights = m_OraConn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY)
            .executeQuery("select distinct weight from ace_frt_adder");
      
      int[] distinctWeights = new int[getResultSetRowCount(weights)];
      for (int i=0; i<distinctWeights.length; i++) {
         weights.next();
         distinctWeights[i] = weights.getInt(1);
      }
      
      return distinctWeights;
   }
   
   private int getResultSetRowCount(ResultSet set) throws Exception {
      int initialRow = set.getRow();
      
      set.last();
      int count = set.getRow();
      
      if (initialRow == 0)
         set.beforeFirst();
      else
         set.absolute(initialRow);
      
      return count; 
   }
   
   private double itemCostToCustomer() throws Exception {
      double
         adjustmentAdderRate = 0.95,
         costToEmery = m_ItemPricingData.getDouble("BUY"),
         adjustedCostToEmery = costToEmery / adjustmentAdderRate,
         
         freightAdderRate = calculateAceFreightAdderRate(m_ItemPricingData.getDouble("SELL"), m_ItemBaseData.getDouble("WEIGHT")),
         freightCost = adjustedCostToEmery * freightAdderRate;
         
         return adjustedCostToEmery + freightCost;
   }
   
   private Map<String, Number> calculateOrgillPricingData() throws Exception {
      double 
         orgillPrice = m_OrgillItemPricingData.getDouble("ORGILL_PRICE"),
         orgillAdder = orgillPrice * ORGILL_ADDER_RATE,
         emeryCustomerPrice = itemCostToCustomer();
      
      orgillPrice += orgillAdder;
      
      if (orgillQuantityRoundAndEmeryBreakableCase()) {
         double orgillBrokenPackCharge = orgillPrice * ORGILL_BROKEN_PACK_CHARGE_RATE;
         orgillPrice += orgillBrokenPackCharge;
      }
         
      double
         emeryOrgillPriceDifference = emeryCustomerPrice - orgillPrice,
         priceDifferencePercentage = emeryOrgillPriceDifference / emeryCustomerPrice;
      
      Map<String, Number> orgillItemPricingData = new HashMap<>();
      orgillItemPricingData.put("ORG_ADDERS", orgillPrice);
      orgillItemPricingData.put("E-O_DIFF", emeryOrgillPriceDifference);
      orgillItemPricingData.put("E-O_PERCENT", priceDifferencePercentage);
      return orgillItemPricingData;
   }
   
   private boolean orgillQuantityRoundAndEmeryBreakableCase() throws Exception {
      return m_OrgillItemPricingData.getBoolean("QTY_ROUND") && !m_OrgillItemPricingData.getBoolean("NBC");
   }

   private void writeAceDispositionColumnEntries() throws Exception {
      writeColumnEntries(m_AceDispositionColumns, m_ItemBaseData);
   }
   
   private void writeDispositionReasonColumnEntries() throws Exception {
      writeColumnEntries(m_ReviewColumns, m_ItemBaseData);
   }

   private String[] loadDispositionReasons() throws Exception {
      ResultSet dispositionReasons = m_SelectDispositionReason.executeQuery();
      dispositionReasons.last();
      int reasonsCount = dispositionReasons.getRow();
      
      String[] reasonsArray = new String[reasonsCount];
      dispositionReasons.beforeFirst();
      for (int i = 0; i < reasonsCount; i++) {
         dispositionReasons.next();
         String reason = dispositionReasons.getString("REASON");
         reasonsArray[i] = reason;
      }
      return reasonsArray;
   }

   private int getNextUnwrittenCellIndex(Row row) {
      int lastCellNumber = row.getLastCellNum();
      return (lastCellNumber == -1) ? 0 : lastCellNumber;
   }
   
   private int getFirstColumnIndex(Object[][] columnGroup) {
      String firstColumnHeading = (String)columnGroup[0][HEADING];
      return getColumnIndexByHeading(firstColumnHeading);
   }
   
   private int getColumnIndexByHeading(String columnHeading) {
      Row headingRow = m_Sheet.getRow(0);
      
      for (int index = -1 ; index < headingRow.getLastCellNum() - 1;) {
         String currentHeading = headingRow.getCell(++index).getStringCellValue();
         if (currentHeading.equals(columnHeading))
            return index;
      }
      return -1;
   }
   
   private void writeCellValue(Object value, Cell cell) {
      if (value instanceof Number) {
         writeCellValue((Number) value, cell);
      } else {
         cell.setCellValue(value.toString());
      }
   }
   
   private void writeCellValue(Number value, Cell cell) {
      if (getStyleType(cell) == STYLE_ID) {
         cell.setCellValue(value.intValue());
      } else {
         cell.setCellValue(value.doubleValue());
      }
   }
   
   private int getStyleType(Cell cell) {
      CellStyle style = cell.getCellStyle();
      return m_Styles.indexOf(style);
   }
   
   // not happy with the length of this method
   private void buildCellStyles() {
      CellStyle
         text = m_Workbook.createCellStyle(),
         id = m_Workbook.createCellStyle(),
         integer = m_Workbook.createCellStyle(),
         decimal = m_Workbook.createCellStyle(),
         money = m_Workbook.createCellStyle(),
         heading = m_Workbook.createCellStyle(),
         percent = m_Workbook.createCellStyle();
      
      text.setAlignment(HorizontalAlignment.LEFT);
      id.setAlignment(HorizontalAlignment.CENTER);
      integer.setAlignment(HorizontalAlignment.RIGHT);
      decimal.setAlignment(HorizontalAlignment.RIGHT);
      money.setAlignment(HorizontalAlignment.RIGHT);
      heading.setAlignment(HorizontalAlignment.CENTER);
      percent.setAlignment(HorizontalAlignment.RIGHT);
      
      DataFormat format = m_Workbook.createDataFormat();
      decimal.setDataFormat(format.getFormat(DECIMAL_FORMAT));
      money.setDataFormat(format.getFormat(MONEY_FORMAT));
      percent.setDataFormat(format.getFormat(PERCENT_FORMAT));
      
      heading.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
      heading.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      heading.setVerticalAlignment(VerticalAlignment.CENTER);
      heading.setWrapText(true);      
      heading.setBorderBottom(BorderStyle.THIN);
      heading.setBorderLeft(BorderStyle.THIN);
      heading.setBorderRight(BorderStyle.THIN);
      
      m_Styles.add(STYLE_TEXT, text);
      m_Styles.add(STYLE_ID, id);
      m_Styles.add(STYLE_INTEGER, integer);
      m_Styles.add(STYLE_FLOAT, decimal);
      m_Styles.add(STYLE_MONEY, money);
      m_Styles.add(STYLE_HEADING, heading);
      m_Styles.add(STYLE_PERCENT, percent);
      
      for (CellStyle style : m_Styles) {
         boolean bold = style.equals(heading);
         style.setFont(buildFont(bold));
      }
   }

   private Font buildFont(boolean isBold) {
      Font font = m_Workbook.createFont();
      
      if (isBold)
         font.setBold(true);
      font.setFontName(FONT_NAME);
      font.setFontHeightInPoints(FONT_SIZE);
      return font;
   }
   
   @Override
   public void setParams(ArrayList<Param> params) {
      String 
         commaSeparatedBuyerCodes = params.get(0).value,
         commaSeparatedDispositions = params.get(1).value;
      m_BuyerCodeList = Arrays.asList(commaSeparatedBuyerCodes.split(","));
      m_DispositionList = Arrays.asList(commaSeparatedDispositions.split(","));
   }
}
