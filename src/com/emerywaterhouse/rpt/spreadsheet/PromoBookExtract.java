/**
 * File: PromoBookExtract.java
 * Description: Promotion Customer Order Book Extract<p>
 *    This is the rewrite of the report so it works with the new report server.
 *    The original author was Peggy Richter.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/16/2005
 * Last Update: $Id: PromoBookExtract.java,v 1.20 2013/02/06 16:44:07 epearson Exp $
 *
 * History
 *    $Log: PromoBookExtract.java,v $
 *    Revision 1.20  2013/02/06 16:44:07  epearson
 *    changed shorts to ints
 *
 *    Revision 1.19  2013/01/16 13:12:06  jfisher
 *    removed oracle specific statement types
 *
 *    Revision 1.18  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.17  2011/11/17 04:27:31  prichter
 *    Changed the way a quantity buy price is obtained in the text version of the report so it matches the Excel version.
 *
 *    Revision 1.16  2011/11/17 00:30:59  prichter
 *    Set the packet id after loading a quantity buy.  Without this, the quantity buy portion of the report doesn't work.
 *
 *    Revision 1.15  2011/06/24 03:35:02  prichter
 *    Added quantity buy columns
 *
 *    Revision 1.14  2011/02/20 07:44:10  prichter
 *    Removed the recalc parameter in getsellprice.
 *
 *    Revision 1.13  2009/05/12 02:56:30  pdavidson
 *    Use dia_date instead of dsb_date when calculating future customer cost.
 *    Per request of Lori and Cetta.
 *
 *    10/20/2005 - Wasn't handling a packet with a null reporting date correctly.   pjr
 *
 *    08/04/2005 - Added parameter to calculate sales history based on packet.report_begin_date.
 *               - Applied changes to run in the report server.  pjr
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    03/17/2005 - Fixed warnings. jcf
 *
 *    06/03/2004 - Pass promotions dsb_date to pricing routine when calculating current sell price
 *              - Added vendor id to report
 *              - Force recalculation of sell price so DPC customers don't pick up a delayed price
 *
 *    05/03/2004 - Removed the setting of the m_DistList variable when sending the report notification.  This variable
 *       will get cleaned up before it can be used by the webservice. - jcf
 *
 *    12/19/2003 - Added meaningful comment so that a blank email is not sent if db connection fails.  pjr
 *
 *    01/08/2003 - Fixed problem with generation of email and account reports not being written to disk
 *
 *    01/07/2003 - Attempted to minimize memory usage by breaking up queries by packet & calling System.gc()
 *
 *    01/02/2003 - Added terms & deadline to customer version - pjr
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.pricing.DiscountLevel;
import com.emerywaterhouse.pricing.DiscountWorksheet;
import com.emerywaterhouse.pricing.WorksheetItem;
import com.emerywaterhouse.pricing.WorksheetLevel;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.utils.StringFormat;
import com.emerywaterhouse.web.QuantityBuy;
import com.emerywaterhouse.websvc.Param;


public class PromoBookExtract extends Report
{
   private final String TAB = "\t";
   private final String NEWLINE = "\r\n";

   // prepared statements
   private PreparedStatement m_CustName;        //gets the customer name
   private PreparedStatement m_Discounts;		//returns the discount id of quantity buys tied to the packet
   private PreparedStatement m_FamilyTree;      //finds all customers connected to a parent id
   private PreparedStatement m_ItemList;        //creates a list of items to report
   private PreparedStatement m_PurchHist;       //units purchased of an item by one or more stores
   private PreparedStatement m_GetSku;          //gets the customer sku for an item
   private PreparedStatement m_PacketInfo;      //returns packet data
   private PreparedStatement m_CurRetail;       //find the customer's current retail for an item
   private PreparedStatement m_CurSell;         //fine the customer's current sell for an item

   //parameters
   private String m_CustId;
   private boolean m_ByAccount;
   private boolean m_ByCustomer;
   private String m_Format;

   private ArrayList<String> m_Packets;
   private String m_AcctFileName;
   private String m_CustFileName;
   private String m_SuppressPacket;
   private String m_AsOfDate;  //pjr 06/29/2005 pass the run date as a parameter for calculating sales history
   private boolean m_UsePacketDate = true;

   //report fields
   private String m_PacketId;
   private String m_Title;
   private String m_Vendor;
   private String m_Message;
   private String m_ItemDescr;
   private String m_ItemId;
   private String m_Sku;
   private String m_Upc;
   private int m_StockPack;
   private String m_Nbc;
   private String m_Unit;
   private double m_CustSell;
   private double m_PromoSell;
   private double m_CustRetail;
   private double m_RetailC;
   private String m_Terms;
   private String m_Deadline;
   private int m_UnitsPurch;
   private int m_VendorId;   // 06/02/2004 - add vendor id to report
   private String m_Warehouse;   // 04/22/2009 - add warehouse to report

   //Quantity buy objects
   private ArrayList<DiscountWorksheet> m_QtyBuyList;

   //report objects
   private StringBuffer m_Lines;
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFCellStyle m_PercentStyle;
   private int m_RowNum = 1;

   // miscellaneous member variables
   private boolean m_Error = false;
   private ArrayList<Integer> m_StoreUnits;
   private ArrayList<String> m_StoreId;
   private int m_MaxRowsPerPage = 65000; // Older Excel files have a row limit of about 65,000
   
   /**
    * default constructor
    */
   public PromoBookExtract()
   {
      super();

      m_Packets = new ArrayList<String>();
      m_StoreUnits = new ArrayList<Integer>();
      m_StoreId = new ArrayList<String>();
      m_QtyBuyList = new ArrayList<DiscountWorksheet>();
   }

   /**
    * Builds a brief description of the break level for an item
    *
    * @param lvl WorksheetItem - the worksheet item
    * @return String - the break level description
    */
   private String buildQBDesc(WorksheetItem item)
   {
   	StringBuffer desc = new StringBuffer();
   	DiscountLevel lvl = item.getDiscountLevel();

   	if ( !lvl.getBreakType().equals("minimum quantity") ) {
   		if ( lvl.getBreakAmount() > 0 ) {
   			desc.append(lvl.getBreakAmount() + " ");

   			if ( lvl.getBreakType().equals("weight") )
   				desc.append("lbs");

   			else
   				if ( lvl.getBreakType().equals("cube") )
   					desc.append("cu ft");

   				else
   					desc.append(lvl.getBreakType());
   		}
   	}

  		if ( item.getMinQuantity() > 0 ) {
  			if ( desc.length() > 0 )
  				desc.append(" ");

  		   desc.append(item.getMinQuantity() + " min qty");
   	}

   	return desc.toString();
   }

   /**
    * Perform cleanup on the objects and close db connections ets.  Overrides the base class
    * method.  The base class method will call closeStatements for us.
    */
   protected void cleanup()
   {
      m_Packets.clear();
      m_StoreUnits.clear();
      m_StoreId.clear();

      DbUtils.closeDbConn(null, m_FamilyTree, null);
      DbUtils.closeDbConn(null, m_CurRetail, null);
      DbUtils.closeDbConn(null, m_CurSell, null);
      DbUtils.closeDbConn(null, m_ItemList, null);
      DbUtils.closeDbConn(null, m_GetSku, null);
      DbUtils.closeDbConn(null, m_PurchHist, null);
      DbUtils.closeDbConn(null, m_Discounts, null);

      m_FamilyTree = null;
      m_CurRetail = null;
      m_CurSell = null;
      m_ItemList = null;
      m_GetSku = null;
      m_PurchHist = null;
      m_Lines = null;
      m_Sheet = null;
      m_WrkBk = null;
      m_CustName = null;
      m_Warehouse = null;
      m_Discounts = null;
      m_QtyBuyList.clear();
      m_QtyBuyList = null;
   }

   /**
    * Runs the report and creates any output that is needed.
    */
   private void closeReport(String reportType, String custid)
   {
      String FileName = "";
      FileOutputStream OutFile = null;

      try {
         if ( reportType.equals("ACCOUNT") )
            FileName = m_AcctFileName;
         else
            FileName = m_CustFileName.replaceFirst("999999", custid);

         FileName = FileName.replaceFirst("ppp", m_PacketId);
         
         if ( m_Format.equals("EXCEL") )
            FileName += ".xlsx";
         
         OutFile = new FileOutputStream(m_FilePath + FileName, false);

         if ( m_Format.equals("EXCEL") ) {
            m_WrkBk.write(OutFile);
         }
         else {
            OutFile.write(m_Lines.toString().getBytes());
            m_Lines.delete(0, m_Lines.length());
         }

         try {
            OutFile.close();
         }
         catch ( Exception e ) {
            log.error("exception", e );
         }

         //
         // Add the file name to the list of files that will be attached or ftp'd
         m_FileNames.add(FileName);
      }
      catch( Exception ex ) {
         log.error("[PromoBookExtract]", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n" + ex.getMessage());
      }

      finally {
         m_Lines = null;
         OutFile = null;
         m_Sheet = null;
         m_WrkBk = null;
      }
   }

   /**
    * Creates a report that shows consolidated sales for the account
    * @param acct
    */
   private void createAcctReport(String acct)
   {
      ResultSet rsStores = null;
      ResultSet items = null;
      int qty;
      int pageNum = 1;
      int itemEaId = 0;

      try {
         //find all stores related to this account and load them into vector m_StoreId
         try {
            m_FamilyTree.setString(1, acct);
            rsStores = m_FamilyTree.executeQuery();
            
            while ( rsStores.next() ) {
               m_StoreId.add( rsStores.getString("customer_id") );
               m_StoreUnits.add(new Integer(0));
            }
         }
         
         catch ( Exception e ) {
            log.error("[PromoBookExtract]", e );
            m_Error = true;
         }
         
         finally {
         	DbUtils.closeDbConn(null, null, rsStores);
            rsStores = null;
         }

         for ( int i = 0; i < m_Packets.size(); i++ ) {
            if ( m_Status != RptServer.RUNNING )
               break;

            setCurAction("Account:" + m_CustId + " packet:" + m_PacketId);

            m_PacketId = m_Packets.get(i);
            m_AsOfDate = getAsOfDate();  //pjr 06/29/2005 Used for sales history date range
            m_ItemList.setString( 1, m_PacketId);
            initReport("ACCOUNT", m_CustId, pageNum);
            items = m_ItemList.executeQuery();

            while ( items.next() && !m_Error && (m_Status == RptServer.RUNNING) ) {
               itemEaId = items.getInt("item_ea_id");
               m_Title = items.getString("title");
               m_Vendor = items.getString("vendor");
               m_Message = items.getString("message");
               m_ItemDescr = items.getString("itemdescr");
               m_ItemId = items.getString("item_id");
               m_Sku = getSku(acct, itemEaId);
               m_Upc = items.getString("upc");
               m_StockPack = items.getInt("stock_pack");
               m_Nbc = items.getString("nbc");
               m_Unit = items.getString("unit");
               m_CustSell = getCurSell(acct, itemEaId, items.getDate("dia_date"));   // 06/02/04 future date base cost. pjr. **PD 5/11/09 changed to dia_date (used to use dsb_date)**
               m_PromoSell = getCurSell(acct, itemEaId, items.getString("promo_id"));
               m_CustRetail = getCurRetail(acct, itemEaId);
               //m_PromoRetail = getCurRetail(acct, m_ItemId, items.getString("promo_id"));
               m_RetailC = items.getDouble("retailc");
               m_Terms = items.getString("terms");
               m_Deadline = items.getString("deadline");
               m_UnitsPurch = 0;
               m_VendorId = items.getInt("vendor_id");  // 06/02/2004 - add vendor id
               m_Warehouse = items.getString("warehouse"); // 04/22/2009 - add warehouse
               
               setCurAction("Account:" + m_CustId + " packet:" + m_PacketId + " vendor:" + m_Vendor + " item:" + m_ItemId);

               for ( int j = 0; j < m_StoreId.size(); j++ ) {
                  qty = unitsSold(m_StoreId.get(j), m_ItemId );
                  m_StoreUnits.set(j, new Integer(qty));
                  m_UnitsPurch = m_UnitsPurch + qty;
               }

               if ( m_Format.equals( "EXCEL" ) ) {
                  if ( m_RowNum > m_MaxRowsPerPage ) {
                     pageNum++;
                     initReport("CUSTOMER", m_CustId, pageNum);
                  }
               }

               createReportLine("ACCOUNT");
            }

            closeReport("ACCOUNT", m_CustId);
            System.gc();
         }
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtract]", e );
         m_Error = true;
      }

      finally {
      	DbUtils.closeDbConn(null, null, rsStores);
      	DbUtils.closeDbConn(null, null, items);
         rsStores = null;
         items = null;
      }
   }

   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return HSSFCell newly created cell
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
    * @return HSSFCell newly created cell
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
    * @return HSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, String val)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));

      return cell;
   }

   private void createCustReport(String cust) throws Exception
   {
      ResultSet items = null;
      int pageNum = 1;
      int itemEaId = 0;

      for ( int i = 0; i < m_Packets.size(); i++ ) {
         if ( m_Status != RptServer.RUNNING ) {
            break;
         }

         try {
            m_PacketId = m_Packets.get( i );
            initReport("CUSTOMER", cust, pageNum);

            m_AsOfDate = getAsOfDate();  //pjr 06/29/2005 Used for sales history date range
            m_ItemList.setString( 1,  m_PacketId );
            items = m_ItemList.executeQuery();
            
            setCurAction("Cust:" + cust + " packet:" + m_PacketId);
            
            while ( items.next() && !m_Error && (m_Status == RptServer.RUNNING) ) {
               itemEaId = items.getInt("item_ea_id");
               
               m_Title = items.getString("title");
               m_Vendor = items.getString("vendor");
               m_Message = items.getString("message");
               m_ItemDescr = items.getString("itemdescr");
               m_ItemId = items.getString("item_id");
               m_Sku = getSku(cust, itemEaId);
               m_Upc = items.getString("upc");
               m_StockPack = items.getInt("stock_pack");
               m_Nbc = items.getString("nbc");
               m_Unit = items.getString("unit");
               m_CustSell = getCurSell(cust, itemEaId, items.getDate("dia_date"));   // 06/02/2004 future date base cost.  pjr **PD 5/11/09 changed to dia_date**
               m_PromoSell = getCurSell(cust, itemEaId, items.getString("promo_id"));
               m_CustRetail = getCurRetail(cust, itemEaId);
               m_RetailC = items.getDouble("retailc");
               m_Terms = items.getString("terms");
               m_Deadline = items.getString("deadline");
               m_UnitsPurch = unitsSold( cust, m_ItemId );
               m_VendorId = items.getInt("vendor_id");  // 06/02/2004 - add vendor id
               m_Warehouse = items.getString("warehouse"); // 04/22/09 - add warehouse
               
               setCurAction("Cust:" + cust + " packet:" + m_PacketId + " vendor:" + m_Vendor + " item:" + m_ItemId);
               createReportLine("CUSTOMER");

               if ( m_Format.equals( "EXCEL" ) ) {
                  if ( m_RowNum > m_MaxRowsPerPage ) {
                     pageNum++;                     
                     initReport("CUSTOMER", cust, pageNum);
                  }
               }
            }

            closeReport("CUSTOMER", cust);
         }

         catch ( Exception e ) {
            log.error("[PromoBookExtract]", e );
            m_Error = true;            
         }

         finally {
         	DbUtils.closeDbConn(null, null, items);
            items = null;
            System.gc();
         }
      }
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      String custId = null;
      ResultSet rs = null;
      m_Status = RptServer.RUNNING;

      try {         
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() ) {
            //
            // If the consolidated account option was chosen, produce that report
            if ( m_ByAccount ) {
               setCurAction("Create account report for " + m_CustId);
               createAcctReport(m_CustId);

               //
               // If the individual reports by customer were requested, produce a
               // report for each customer within this account
               if ( m_ByCustomer ) {
                  m_FamilyTree.setString(1, m_CustId);
                  rs = m_FamilyTree.executeQuery();

                  while ( rs.next() && m_Status == RptServer.RUNNING) {
                     custId = rs.getString("customer_id");
                     setCurAction("Create customer report for " + custId);
                     createCustReport(custId);
                  }
               }
            }
            else {
               //
               // if only a single customer report was requested, produce that report
               if ( m_ByCustomer ) {
                  setCurAction("Create customer report for " + m_CustId);                  
                  createCustReport(m_CustId);                  
               }
            }

            created = true;
         }
      }

      catch ( Exception ex ) {
         log.fatal("[PromoBookExtract]", ex);
         m_ErrMsg.append("The remaining reports experienced errors and were not generated" + NEWLINE + NEWLINE);
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);
         rs = null;
         cleanup();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    *
    * @param reportType
    */
   private void createReportLine(String reportType)
   {
      int col;
      XSSFRow Row = null;
      XSSFCell cell = null;
      Integer qty;
      DiscountWorksheet ws = null;
      WorksheetLevel lvl = null;
      WorksheetItem item = null;      
      boolean qbFound = false;
      double price = 0.0;

      if ( m_Format.equals( "EXCEL" ) ){
         m_RowNum++;
         Row = m_Sheet.createRow(m_RowNum);

         if ( Row != null ) {
            col = 0;

            //only print the packet id if it isn't in the excluded packet list
            if ( m_SuppressPacket.indexOf( m_PacketId ) == -1 ) {
               createCell(Row, col++, m_PacketId);
            }
            
            else {
               createCell(Row, col++, " ");
            }
            
            createCell(Row, col++, m_Title);
            createCell(Row, col++, m_VendorId);  // 06/02/2004 - add vendor id. pjr
            createCell(Row, col++, m_Vendor);
            createCell(Row, col++, m_Message);
            createCell(Row, col++, m_ItemDescr);
            createCell(Row, col++, m_ItemId);
            createCell(Row, col++, m_Sku);
            createCell(Row, col++, m_Upc);
            createCell(Row, col++, m_StockPack);
            createCell(Row, col++, m_Nbc);
            createCell(Row, col++, m_Unit);
            createCell(Row, col++, m_CustSell);
            createCell(Row, col++, m_PromoSell);

            // savings displayed as percent
            cell = createCell(Row, col++, (m_CustSell - m_PromoSell) / m_CustSell);
            cell.setCellStyle(m_PercentStyle);

            // If a quantity buy is tied to this packet, show
            // the available discounts for this item
            if ( m_QtyBuyList.size() > 0 ) {
            	for ( int i = 0; i < m_QtyBuyList.size(); i++ ) {
            		ws = m_QtyBuyList.get(i);

            		for ( int j = 0; j < ws.getDiscountLevelCount(); j++ ) {
            			lvl = ws.getDiscountLevel(j);

            			for ( int k = 0; k < lvl.getItemCount(); k++ ) {
            				item = lvl.getItem(k);
            				qbFound = false;

            				if ( item.getItemId().equals(m_ItemId) ) {
            					price = item.calcNewPrice();
            					createCell(Row, col++, buildQBDesc(item));
            					createCell(Row, col++, price);
            					cell = createCell(Row, col++, (m_CustSell - price) / m_CustSell);
            					cell.setCellStyle(m_PercentStyle);
            					qbFound = true;
            					break;
            				}
            			}

            			// If an item was not printed, advance the column count past
            			// this discount level
            			if ( !qbFound ) {
           					col += 3;
            			}
            		}
            	}
            }

            createCell(Row, col++, m_CustRetail);
            createCell(Row, col++, m_RetailC);
            createCell(Row, col++, m_Terms);
            createCell(Row, col++, m_Deadline);
            createCell(Row, col++, m_Warehouse);
            createCell(Row, col++, m_UnitsPurch);

            if ( reportType.equals("ACCOUNT") ) {
               for ( int i = 0; i < m_StoreId.size(); i++ ) {
                  qty = m_StoreUnits.get(i);
                  createCell(Row, col++, qty.intValue());
               }
            }
            else {
               createCell(Row, col++, "__________");
            }
        }
      }

      else {
         //only print the packet id if it isn't in the excluded packet list
         if ( m_SuppressPacket.indexOf( m_PacketId ) == -1 ) {
            m_Lines.append(m_PacketId + TAB);
         }
         
         else {
            m_Lines.append(TAB);
         }

         m_Lines.append(m_Title + TAB);
         m_Lines.append(m_VendorId + TAB);  // 06/02/2004 - add vendor id. pjr
         m_Lines.append(m_Vendor + TAB);
         m_Lines.append(m_Message + TAB);
         m_Lines.append(m_ItemDescr + TAB);
         m_Lines.append(m_ItemId + TAB);
         m_Lines.append(m_Sku + TAB);
         m_Lines.append(m_Upc + TAB);
         m_Lines.append(m_StockPack + TAB);
         m_Lines.append(m_Nbc + TAB);
         m_Lines.append(m_Unit + TAB);
         m_Lines.append(m_CustSell + TAB);
         m_Lines.append(m_PromoSell + TAB);
         m_Lines.append((m_CustSell - m_PromoSell) / m_CustSell * 100 + TAB);

         // If a quantity buy is tied to this packet, show
         // the available discounts for this item
         if ( m_QtyBuyList.size() > 0 ) {
         	for ( int i = 0; i < m_QtyBuyList.size(); i++ ) {
         		ws = m_QtyBuyList.get(i);

         		for ( int j = 0; j < ws.getDiscountLevelCount(); j++ ) {
         			lvl = ws.getDiscountLevel(j);

         			for ( int k = 0; k < lvl.getItemCount(); k++ ) {
         				item = lvl.getItem(k);
         				qbFound = false;

         				if ( item.getItemId().equals(m_ItemId) ) {
         					price = item.calcNewPrice();
         					m_Lines.append(buildQBDesc(item) + TAB);
         					m_Lines.append(price + TAB);
         					m_Lines.append((m_CustSell - price) / m_CustSell * 100);
         					qbFound = true;
         					break;
         				}
         			}

         			// If an item was not printed, advance the column count past
         			// this discount level
         			if ( !qbFound )
         				m_Lines.append("\t\t\t");
         		}
         	}
         }

         m_Lines.append(m_CustRetail + TAB);
         m_Lines.append(m_RetailC + TAB);
         m_Lines.append(m_Terms + TAB);
         m_Lines.append(m_Deadline + TAB);
         m_Lines.append(m_Warehouse + TAB);
         m_Lines.append(m_UnitsPurch + TAB);

         if ( reportType.equals("ACCOUNT") ) {
            for ( int i = 0; i < m_StoreId.size(); i++ ) {
               m_Lines.append( m_StoreUnits.get(i) + TAB );
            }
         }

         m_Lines.append( NEWLINE );
      }
   }

   /**
    * Returns the correct sales history end date for the current packet
    */
   private String getAsOfDate()
   {
      String dateStr = null;
      ResultSet rs = null;

      if ( !m_UsePacketDate )
         return m_AsOfDate;

      try {
         m_PacketInfo.setString(1, m_PacketId);
         rs = m_PacketInfo.executeQuery();

         if ( rs.next() )
            dateStr = rs.getString("RepDate");
         else {
            log.error("Unable to retrieve date for packet " + m_PacketId);
            m_Error = true;
         }
      }

      catch ( Exception e ) {
         log.error("exception", e );
         m_Error = true;
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return dateStr;
   }

   /**
    * overrides base class method for logging.
    * @return The id of the customer from the params passed to the report.
    * @see com.emerywaterhouse.rpt.server.Report#getCustId()
    */
   @Override
   public String getCustId()
   {
      return m_CustId;
   }

   /**
    * Gets the name of the customer.
    * @param custid
    * @return the customer name
    */
   private String getCustName(String custid)
   {
      ResultSet rs = null;
      String name = null;

      try {
         m_CustName.setString(1, custid);
         rs = m_CustName.executeQuery();

         while ( rs.next() )
            name = rs.getString("name");
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtract]", e );
         m_Error = true;
      }

      finally {
      	 DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return name;
   }
   
   /**
    * @return the maxRowsPerPage
    */
   public int getMaxRowsPerPage() 
   {
      return m_MaxRowsPerPage;
   }

   /**
    * Attempt to find the current retail for an item
    *
    * @param custid
    * @param itemid
    * @return double 
    */
   private double getCurRetail(String custid, int itemEaId)
   {
      ResultSet rs = null;
      double retail = 0;
      
      try {
         m_CurRetail.setString(1, custid);
         m_CurRetail.setInt(2, itemEaId);
         m_CurRetail.setNull(3, Types.VARCHAR);
         
         rs = m_CurRetail.executeQuery();
         
         if ( rs.next() )
            retail = rs.getDouble(1);
      }
            
      catch (Exception e) {
         log.error("[PromoBookExtract]", e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      return retail;
   }
   
   /**
    * Attempt to find the current sell for an item
    *
    * @param custid
    * @param itemid
    * @return float
    *
    * 5/11/09 - Now uses dia_date instead of dsb_date PD
    * 06/02/2004 - Pass promo's dsb_date to pricing routine when calculating current sell
    */
   private double getCurSell(String custid, int itemEaId, java.sql.Date asOf)
   {
      ResultSet rs = null;
      double sell = 0;
      
      try {
         m_CurSell.setString(1, custid);
         m_CurSell.setInt(2, itemEaId);
         m_CurSell.setNull(3, Types.VARCHAR);
         m_CurSell.setDate(4, asOf);

         rs = m_CurSell.executeQuery();
         
         if ( rs.next() )
            sell = rs.getDouble(1);
      }
            
      catch (Exception e) {
         log.error("[PromoBookExtract]", e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      return sell;
   }

   /**
    * Returns the promotional sell price of an item
    *
    * @param custid
    * @param itemid
    * @param promoid
    * @return the current sell price
    */
   private double getCurSell(String custid, int itemEaId, String promoid)
   {
      ResultSet rs = null;
      double sell = 0;
      
      try {
         m_CurSell.setString(1, custid);
         m_CurSell.setInt(2, itemEaId);
         m_CurSell.setString(3, promoid);
         m_CurSell.setNull(4, Types.DATE);
         m_CurSell.execute();
         
         rs = m_CurSell.executeQuery();
         
         if ( rs.next() )
            sell = rs.getDouble(1);
      }
            
      catch (Exception e) {
         log.error("[PromoBookExtract]", e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      return sell;
   }

   /**
    * Returns a customer sku for a given item
    *
    * @param custid - the customer id
    * @param itemid - the item id
    * @return String - the customer sku
    */
   private String getSku(String custid, int itemEaId)
   {
      ResultSet rs = null;
      String sku =  " ";

      try {
         m_GetSku.setString(1, custid);
         m_GetSku.setInt(2, itemEaId);
         rs = m_GetSku.executeQuery();

         if ( rs.next() )
            sku = rs.getString("customer_sku");
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtract]", e );
         m_Error = true;
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return sku;
   }

   /**
    * Creates the report headings
    *
    * @param reportType
    * @param custid
    */
   private void initReport(String reportType, String custid, int pageNum)
   {
      XSSFRow Row = null;
      int col;
      ResultSet rs = null;
      int discountId;
      QuantityBuy qtyBuy = null;
      DiscountWorksheet ws = null;
      WorksheetLevel lvl = null;

      try {
         if ( pageNum == 1 ) {
            m_QtyBuyList.clear();
            m_Discounts.setString(1, m_PacketId);
            rs = m_Discounts.executeQuery();

            while ( rs.next() ) {
            	discountId = rs.getInt("discount_id");

            	if ( qtyBuy == null )
            		qtyBuy = new QuantityBuy(m_EdbConn);

            	qtyBuy.setPacketId(m_PacketId);
            	            	            	
            	ws = qtyBuy.loadDiscountWorksheet(discountId, m_CustId);

            	if ( ws != null ) {
            		m_QtyBuyList.add(ws);
            	}
            }
         }

         if ( m_Format.equals("EXCEL") ) {
            if ( pageNum == 1 ) {
               m_WrkBk = new XSSFWorkbook();
               m_PercentStyle = m_WrkBk.createCellStyle();
               m_PercentStyle.setDataFormat(m_WrkBk.createDataFormat().getFormat("0.0%"));
            }

            m_Sheet = m_WrkBk.createSheet(String.format("Page %d", pageNum));

            //display the customer id and name
            Row = m_Sheet.createRow(0);
            if ( Row != null ) {
               if ( reportType.equals("ACCOUNT"))
                  createCell(Row, 0, "Account");
               else
                  createCell(Row, 0, "Customer");

               createCell(Row, 1, custid);
               createCell(Row, 2, getCustName(custid));
            }

            Row = m_Sheet.createRow(1);
            m_RowNum = 1;
            col = 0;

            if ( Row != null ) {
               createCell(Row, col++, "Packet");
               createCell(Row, col++, "Title");
               createCell(Row, col++, "Vendor Id");  //06/02/2004 - add vendor name. pjr
               createCell(Row, col++, "Vendor Name");
               createCell(Row, col++, "Message");
               createCell(Row, col++, "Item Description");
               createCell(Row, col++, "Item Id");
               createCell(Row, col++, "SKU");
               createCell(Row, col++, "UPC");
               createCell(Row, col++, "Stock Pack");
               createCell(Row, col++, "NBC");
               createCell(Row, col++, "Ship Unit");
               createCell(Row, col++, "Cost");
               createCell(Row, col++, "Promo Cost");
               createCell(Row, col++, "%Saved");

               //
               // If a quantity buy is tied to this packet, insert columns to show the break points
               if ( m_QtyBuyList.size() > 0 ) {
                  for ( int i = 0; i < m_QtyBuyList.size(); i++ ) {
                     ws = m_QtyBuyList.get(i);

                     for ( int j = 0; j < ws.getDiscountLevelCount(); j++ ) {
                        lvl = ws.getDiscountLevel(j);
                        createCell(Row, col++, lvl.getLevelDescr());
                        createCell(Row, col++, "QB Cost");
                        createCell(Row, col++, "%Saved");
                     }
                  }
               }

               createCell(Row, col++, "Retail");
               createCell(Row, col++, "Retail C");
               createCell(Row, col++, "Terms");
               createCell(Row, col++, "Deadline");
               createCell(Row, col++, "Warehouse");
               createCell(Row, col++, "Units Ordered");

               if ( reportType.equals("ACCOUNT") ) {
                  for ( int i = 0; i < m_StoreId.size(); i++ ) {
                     createCell(Row, col++, m_StoreId.get(i));
                  }
               }
               else
                  createCell(Row, col++, "Order Qty");
            }
         }
         else {
            m_Lines = new StringBuffer();

            if ( reportType.equals("ACCOUNT") )
               m_Lines.append("Account" + "\t" + custid + "\t" + getCustName(custid) + "\r\n");
            else
               m_Lines.append("Customer" + "\t" + custid + "\t" + getCustName(custid) + "\r\n");

            // 06/02/2004 - Add vendor id. pjr
            m_Lines.append("Packet\tTitle\tVendor Id\tVendor Name\tMessage\tItem Description\tItem Id\tSKU\tUPC\tStock Pack\tNBC\tShip Unit\t");
            m_Lines.append("Cost\tPromo Cost\t%Saved\t");

            // If a quantity buy is tied to this packet, insert
            // columns to show the break points
            if ( m_QtyBuyList.size() > 0 ) {
            	for ( int i = 0; i < m_QtyBuyList.size(); i++ ) {
            		ws = m_QtyBuyList.get(i);

            		for ( int j = 0; j < ws.getDiscountLevelCount(); j++ ) {
            			lvl = ws.getDiscountLevel(j);
            			m_Lines.append(lvl.getLevelDescr());
            			m_Lines.append("\tQB Cost\t%Saved\t");
            		}
            	}
            }

            m_Lines.append("Retail\tRetail C\t");
            m_Lines.append("Terms\tDeadline\tWarehouse\tUnits Ordered\t");

            if ( reportType.equals("ACCOUNT") ) {
               for ( int i = 0; i < m_StoreId.size(); i++ ) {
                  m_Lines.append(m_StoreId.get(i) + "\t" );
               }
            }
            else {
               m_Lines.append("Order Qty");
            }

            m_Lines.append( NEWLINE );
         }
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtract]", e );
         m_Error = true;
      }

      finally {
         Row = null;
      }
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @return boolean
    * @throws Exception
    */
   private boolean prepareStatements() throws Exception
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();

      if ( m_EdbConn != null ) {
         try {
            m_CustName = m_EdbConn.prepareStatement("select name from customer where customer_id = ?");

            sql.setLength(0);
            sql.append("select distinct quantity_buy.discount_id ");
            sql.append("from promotion ");
            sql.append("join quantity_buy on quantity_buy.packet_id = promotion.packet_id or quantity_buy.promo_id = promotion.promo_id ");
            sql.append("join discount on discount.discount_id = quantity_buy.discount_id and ");
            sql.append("     beg_date <= promotion.po_begin and (end_date is null or end_date >= promotion.po_begin) ");
            sql.append("join discount_break_type on discount_break_type.break_type_id = quantity_buy.break_type_id ");
            sql.append("where promotion.packet_id = ? ");
            m_Discounts = m_EdbConn.prepareStatement(sql.toString());

            // 06/02/2004 - optionally pass dsb date to pricing routine
            m_CurSell = m_EdbConn.prepareStatement("select round(price, 2) as sell from ejd_cust_procs.get_sell_price(?, ?, ?, ?)");
            m_CurRetail = m_EdbConn.prepareStatement("select ejd_price_procs.get_retail_price(?, ?, ?)");  //custid, ieaid, promoid
            
            sql.setLength(0);
            sql.append("select name, customer_id ");
            sql.append("from customer ");
            sql.append("start with customer_id = ? connect by prior customer_id = parent_id");
            m_FamilyTree = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select ");
            sql.append("packet.packet_id, ");
            sql.append("packet.title, ");
            sql.append("iea.item_id, ");
            sql.append("iea.item_ea_id, iea.description as itemdescr, ");
            sql.append("vendor.name as vendor, ");
            sql.append("nvl(ppi.message, ' ') as message, ");
            sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', ' ', 'N') as nbc, ");
            sql.append("eiw.stock_pack, ");
            sql.append("ship_unit.unit, ");
            sql.append("nvl(upc_code, ' ') as upc, ");
            sql.append("promo_item.promo_id, ");
            sql.append("terms.name as terms, ");
            sql.append("to_char(dia_date, 'mm/dd/yyyy') as deadline, ");
            sql.append("retail_c as retailc, ");
            sql.append("dia_date, ");
            sql.append("vendor.vendor_id, ");
            sql.append("(SELECT string_agg(w2.name, ', ') ");
            sql.append(" FROM ejd_item_warehouse ");
            sql.append(" JOIN warehouse w2 ON ejd_item_warehouse.warehouse_id = w2.warehouse_id ");
            sql.append(" WHERE ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id AND ejd_item_warehouse.warehouse_id = promotion.warehouse_id ");
            sql.append(" GROUP BY ejd_item_id) AS warehouse ");
            sql.append("from packet ");
            sql.append("join promotion on promotion.packet_id = packet.packet_id ");
            sql.append("join promo_item on promo_item.promo_id = promotion.promo_id ");
            sql.append("join item_entity_attr iea on iea.item_ea_id = promo_item.item_ea_id ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = iea.ejd_item_id ");
            sql.append("join ejd_item_warehouse eiw on ejd_item.ejd_item_id = eiw.ejd_item_id and eiw.warehouse_id = promotion.warehouse_id ");
            sql.append("left outer join ejd_item_whs_upc eiwu on eiwu.ejd_item_id = ejd_item.ejd_item_id and eiwu.warehouse_id = promotion.warehouse_id and primary_upc = 1 ");
            sql.append("join ejd_item_price eip on eip.ejd_item_id = ejd_item.ejd_item_id and eip.warehouse_id = promotion.warehouse_id ");
            sql.append("left outer join preprint_item ppi on ppi.promo_item_id = promo_item.promo_item_id ");
            sql.append("join broken_case on  broken_case.broken_case_id = ejd_item.broken_case_id ");
            sql.append("join ship_unit on ship_unit.unit_id = iea.ship_unit_id ");
            sql.append("join terms on terms.term_id = promotion.term_id ");
            sql.append("join vendor on vendor.vendor_id = iea.vendor_id ");
            sql.append("where packet.packet_id = ? ");
            sql.append("order by vendor.name, iea.description ");
            
            m_ItemList = m_EdbConn.prepareStatement(sql.toString());
            sql.setLength(0);

            m_PacketInfo = m_EdbConn.prepareStatement(
               "select to_char(nvl(report_begin_date, current_date), 'mm/dd/yyyy') repdate from packet where packet_id = ?"
            );

            // pjr 06/29/2005 - pass the end of the date range
            sql.setLength(0);
            sql.append("select sum(qty_shipped) qty from inv_dtl ");
            sql.append("where cust_nbr = ? and item_nbr = ? and ");
            sql.append("invoice_date <= current_date and invoice_date >= current_date - 365");
            m_PurchHist = m_EdbConn.prepareStatement(sql.toString());

            m_GetSku = m_EdbConn.prepareStatement(
               "select customer_sku from item_ea_cross where customer_id = ejd_cust_procs.get_cons_id(?, 'ITEM XREF') and item_ea_id = ?"
            );

            isPrepared = true;
         }

         catch ( Exception ex ) {
            log.fatal("[PromoBookExtract]", ex);
         }
      }

      return isPrepared;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Note - m_Email, and m_Zipped have been removed from the params the report gets.
    */
   public void setParams(ArrayList<Param> params)
   {
      m_CustId = params.get(0).value;

      m_ByCustomer = Boolean.parseBoolean(params.get(1).value);
      m_ByAccount = Boolean.parseBoolean(params.get(2).value);
      m_Format = params.get(3).value;
      m_SuppressPacket = params.get(4).value;
      m_AcctFileName = params.get(5).value;
      m_CustFileName = params.get(6).value;
      m_Packets = StringFormat.parseString(params.get(7).value, ';');

      //pjr 06/29/2005 New parameter.  Either a report date or 'default'
      m_AsOfDate = params.get(8).value;
      m_UsePacketDate = ( m_AsOfDate.equals("default") );
   }
      
   /**
    * @param maxRowsPerPage the maxRowsPerPage to set
    */
   public void setMaxRowsPerPage(int maxRowsPerPage) {
      this.m_MaxRowsPerPage = maxRowsPerPage;
   }

   /**
    *
    * @param custid
    * @param itemid
    * @return the number of units sold
    */
   private int unitsSold(String custid, String itemid)
   {
      ResultSet rs = null;
      int qty = 0;

      try{
         m_PurchHist.setString(1, custid);
         m_PurchHist.setString(2, itemid);
        
         rs = m_PurchHist.executeQuery();

         if ( rs.next() )
            qty = rs.getInt("qty");
      }
      catch ( Exception e ) {
         log.error("[PromoBookExtract]",  e );
         m_Error = true;
      }
      finally {
      	DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return qty;
   }
}