package com.emerywaterhouse.rpt.export;

/**
 * File: AceTool.java
 * Description: catalog export for Ace Tool.  They can't handle parsing any data themselves
 *    and need us to create an export that they can use to add items to an off the shelf
 *    CMS application.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 05/18/2010
 * Last Update: $Id: AceTool.java,v 1.3 2012/07/11 17:30:56 jfisher Exp $
 *
 * History
 *    $Log: AceTool.java,v $
 *    Revision 1.3  2012/07/11 17:30:56  jfisher
 *    in_catalog modification
 *
 *    Revision 1.2  2012/07/11 17:11:28  jfisher
 *    in_catalog modification
 *
 *    Revision 1.1  2010/05/19 19:26:48  jfisher
 *    Initial add
 *
 */

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class AceTool extends Report
{
   private static final String aceCustId = "016829";
   private static final short MAX_COLS = 93;

   //
   // IDs for the styles used in the worksheet.
   private static final int stText     = 0;
   private static final int stInt      = 1;
   private static final int stDouble2  = 2;
   private static final int stDouble3  = 3;

   private PreparedStatement m_ItemData;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private HSSFCellStyle[] m_CellStyles;

   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;

   /**
    *
    */
   public AceTool()
   {
      super();
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      setupWorkbook();
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < m_CellStyles.length; i++ )
            m_CellStyles[i] = null;
      }

      m_Sheet = null;
      m_Wrkbk = null;
      m_CellStyles = null;

      super.finalize();
   }

   /**
    * adds a double type cell to current row at the specified column in current sheet
    *
    * @param row The row that the cell will be added to.
    * @param col 0-based column number of spreadsheet cell
    * @param style ID of the Excel style to be used to display cell
    * @param value double value to be stored in cell
    */
   private void addCell(HSSFRow row, int col, int style, double value)
   {
      HSSFCell cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(value);
      cell.setCellStyle(m_CellStyles[style]);
      cell = null;
   }

   /**
    * adds an integer type cell to current row at the specified column in current sheet
    *
    * @param row The row that the cell will be added to.
    * @param col 0-based column number of spreadsheet cell
    * @param style ID of the Excel style to be used to display cell
    * @param value integer value to be stored in cell
    */
   private void addCell(HSSFRow row, int col, int style, int value)
   {
      HSSFCell cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(value);
      cell.setCellStyle(m_CellStyles[style]);
      cell = null;
   }

   /**
    * adds a text type cell to current row at the specified column in current sheet
    *
    * @param row The row that the cell will be added to.
    * @param col 0-based column number of spreadsheet cell
    * @param style ID of the Excel style to be used to display cell
    * @param value text value to be stored in cell
    */
   private void addCell(HSSFRow row, int col, int style, String value)
   {
      HSSFCell cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new HSSFRichTextString(value));
      cell.setCellStyle(m_CellStyles[style]);
      cell = null;
   }

   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      HSSFRow row = null;
      int rowNum = 0;
      FileOutputStream outFile = null;
      ResultSet itemData = null;
      boolean result = false;
      String itemId = null;
      String desc = null;
      String upc = null;
      String vndName = null;
      String largeImg = null;
      String smallImg = null;
      String thumbImg = null;
      double weight = 0.0;
      double cost = 0.0;
      double retail = 0.0;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         rowNum = createCaptions();

         m_ItemData.setString(1, aceCustId);
         m_ItemData.setString(2, aceCustId);
         m_ItemData.setString(3, aceCustId);
         itemData = m_ItemData.executeQuery();

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            itemId = itemData.getString("item_id");
            desc = itemData.getString("description");
            upc = itemData.getString("upc_code");
            vndName = itemData.getString("vnd_name");
            weight = itemData.getDouble("weight");
            cost = itemData.getDouble("cost");
            retail = itemData.getDouble("retail");
            smallImg = itemData.getString("img_md");
            thumbImg = itemData.getString("img_sm");
            largeImg = itemData.getString("img_lg");

            setCurAction("processing item: " + itemId);

            row = createRow(rowNum++);

            if ( row != null ) {
               for ( int i = 0; i < MAX_COLS; i++ ) {
                  switch( i ) {
                     case 0: {
                        addCell(row, i, stText, "default"); //a
                        break;
                     }

                     case 1: {
                        addCell(row, i, stText, "base"); //b
                        break;
                     }

                     case 2: {
                        addCell(row, i, stText, "default"); //c
                        break;
                     }

                     case 3: {
                        addCell(row, i, stText, "simple"); //d
                        break;
                     }

                     case 4: {
                        addCell(row, i, stInt, 2); //e
                        break;
                     }

                     case 5: {
                        addCell(row, i, stText, itemId); //f
                        break;
                     }

                     case 6: {
                        addCell(row, i, stInt, 0); //g
                        break;
                     }

                     case 7: {
                        addCell(row, i, stText, desc); //h
                        break;
                     }

                     case 8: {
                        addCell(row, i, stText, desc); //i
                        break;
                     }

                     case 9: {
                        addCell(row, i, stText, desc); //j
                        break;
                     }

                     case 10: {
                        addCell(row, i, stText, largeImg); //k
                        break;
                     }

                     case 11: {
                        addCell(row, i, stText, smallImg); //l
                        break;
                     }

                     case 12: {
                        addCell(row, i, stText, thumbImg); //m
                        break;
                     }

                     case 13: {
                        addCell(row, i, stText, ""); //n
                        break;
                     }

                     case 14: {
                        addCell(row, i, stText, ""); //o
                        break;
                     }

                     case 15: {
                        addCell(row, i, stText, ""); //p
                        break;
                     }

                     case 16: {
                        addCell(row, i, stText, 2); //q
                        break;
                     }

                     case 17: {
                        addCell(row, i, stText, "Product Info Column"); //r
                        break;
                     }

                     case 18: {
                        addCell(row, i, stText, ""); //s
                        break;
                     }

                     case 19: {
                        addCell(row, i, stDouble2, retail); //t
                        break;
                     }

                     case 20: {
                        addCell(row, i, stText, ""); //u
                        break;
                     }

                     case 21: {
                        addCell(row, i, stDouble3, cost); //v
                        break;
                     }

                     case 22: {
                        addCell(row, i, stDouble2, weight); //w
                        break;
                     }

                     case 23: {
                        addCell(row, i, stText, desc); //x
                        break;
                     }

                     case 24: {
                        addCell(row, i, stText, desc); //y
                        break;
                     }

                     case 25: {
                        addCell(row, i, stText, ""); //z
                        break;
                     }

                     case 26: {
                        addCell(row, i, stText, ""); //aa
                        break;
                     }

                     case 27: {
                        addCell(row, i, stText, vndName); //ab
                        break;
                     }

                     case 28: {
                        addCell(row, i, stText, "Enabled"); //ac
                        break;
                     }

                     case 29: {
                        addCell(row, i, stText, "Taxable Goods"); //ad
                        break;
                     }

                     case 30: {
                        addCell(row, i, stText, "Catalog, Search"); //ae
                        break;
                     }

                     case 31: {
                        addCell(row, i, stText, "Yes"); //af
                        break;
                     }

                     case 32: {
                        addCell(row, i, stText, ""); //ag
                        break;
                     }

                     case 33: {
                        addCell(row, i, stText, ""); //ah
                        break;
                     }

                     case 34: {
                        addCell(row, i, stText, ""); //ai
                        break;
                     }

                     case 35: {
                        addCell(row, i, stText, ""); //aj
                        break;
                     }

                     case 36: {
                        addCell(row, i, stText, ""); //ak
                        break;
                     }

                     case 37: {
                        addCell(row, i, stText, ""); //al
                        break;
                     }

                     case 38: {
                        addCell(row, i, stText, ""); //am
                        break;
                     }

                     case 39: {
                        addCell(row, i, stText, ""); //an
                        break;
                     }

                     case 40: {
                        addCell(row, i, stText, ""); //ao
                        break;
                     }

                     case 41: {
                        addCell(row, i, stText, ""); //ap
                        break;
                     }

                     case 42: {
                        addCell(row, i, stText, ""); //aq
                        break;
                     }

                     case 43: {
                        addCell(row, i, stText, ""); //ar
                        break;
                     }

                     case 44: {
                        addCell(row, i, stInt, 0); //as
                        break;
                     }

                     case 45: {
                        addCell(row, i, stInt, 1); //at
                        break;
                     }

                     case 46: {
                        addCell(row, i, stInt, 0); //au
                        break;
                     }

                     case 47: {
                        addCell(row, i, stInt, 1); //av
                        break;
                     }

                     case 48: {
                        addCell(row, i, stText, ""); //aw
                        break;
                     }

                     case 49: {
                        addCell(row, i, stText, ""); //ax
                        break;
                     }

                     case 50: {
                        addCell(row, i, stText, ""); //ay
                        break;
                     }

                     case 51: {
                        addCell(row, i, stInt, 0); //az
                        break;
                     }

                     case 52: {
                        addCell(row, i, stInt, 1); //ba
                        break;
                     }

                     case 53: {
                        addCell(row, i, stText, ""); //bb
                        break;
                     }

                     case 54: {
                        addCell(row, i, stText, ""); //bc
                        break;
                     }

                     case 55: {
                        addCell(row, i, stText, ""); //bd
                        break;
                     }

                     case 56: {
                        addCell(row, i, stInt, 1); //be
                        break;
                     }

                     case 57: {
                        addCell(row, i, stInt, 1); //bf
                        break;
                     }

                     case 58: {
                        addCell(row, i, stText, ""); //bg
                        break;
                     }

                     case 59: {
                        addCell(row, i, stInt, 0); //bh
                        break;
                     }

                     case 60: {
                        addCell(row, i, stInt, 1); //bi
                        break;
                     }

                     case 61: {
                        addCell(row, i, stInt, 1); //bj
                        break;
                     }

                     case 62: {
                        addCell(row, i, stText, ""); //bk
                        break;
                     }

                     case 63: {
                        addCell(row, i, stText, ""); //bl
                        break;
                     }

                     case 64: {
                        addCell(row, i, stText, desc); //bm
                        break;
                     }

                     case 65: {
                        addCell(row, i, stText, desc); //bn
                        break;
                     }

                     case 66: {
                        addCell(row, i, stText, desc); //bo
                        break;
                     }

                     case 67: {
                        addCell(row, i, stText, ""); //bp
                        break;
                     }

                     case 68: {
                        addCell(row, i, stText, ""); //bq
                        break;
                     }

                     case 69: {
                        addCell(row, i, stText, ""); //br
                        break;
                     }

                     case 70: {
                        addCell(row, i, stText, ""); //bs
                        break;
                     }

                     case 71: {
                        addCell(row, i, stText, ""); //bt
                        break;
                     }

                     case 72: {
                        addCell(row, i, stText, ""); //bu
                        break;
                     }

                     case 73: {
                        addCell(row, i, stText, ""); //bv
                        break;
                     }

                     case 74: {
                        addCell(row, i, stText, ""); //bw
                        break;
                     }

                     case 75: {
                        addCell(row, i, stText, ""); //bx
                        break;
                     }

                     case 76: {
                        addCell(row, i, stText, ""); //by
                        break;
                     }

                     case 77: {
                        addCell(row, i, stText, ""); //bz
                        break;
                     }

                     case 78: {
                        addCell(row, i, stText, ""); //ca
                        break;
                     }

                     case 79: {
                        addCell(row, i, stText, ""); //cb
                        break;
                     }

                     case 80: {
                        addCell(row, i, stText, upc); //cc
                        break;
                     }

                     case 81: {
                        addCell(row, i, stText, ""); //cd
                        break;
                     }

                     case 82: {
                        addCell(row, i, stText, ""); //ce
                        break;
                     }

                     case 83: {
                        addCell(row, i, stText, ""); //cf
                        break;
                     }

                     case 84: {
                        addCell(row, i, stText, ""); //cg
                        break;
                     }

                     case 85: {
                        addCell(row, i, stText, ""); //cg
                        break;
                     }

                     case 86: {
                        addCell(row, i, stText, ""); //ci
                        break;
                     }

                     case 87: {
                        addCell(row, i, stText, ""); //cj
                        break;
                     }

                     case 88: {
                        addCell(row, i, stText, ""); //ck
                        break;
                     }

                     case 89: {
                        addCell(row, i, stText, ""); //cl
                        break;
                     }

                     case 90: {
                        addCell(row, i, stText, ""); //cm
                        break;
                     }

                     case 91: {
                        addCell(row, i, stText, ""); //cn
                        break;
                     }

                     case 92: {
                        addCell(row, i, stText, ""); //co
                        break;
                     }
                  }
               }
            }
         }

         m_Wrkbk.write(outFile);
         itemData.close();

         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         row = null;
         outFile = null;
         itemId = null;
         desc = null;
         upc = null;
         vndName = null;
         largeImg = null;
         smallImg = null;
         thumbImg = null;
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemData);
   }

   /**
    * Sets the captions on the report.
    */
   private short createCaptions()
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      short rowNum = 0;

      if ( m_Sheet == null )
         return 0;

      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum);

      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            cell = row.createCell(i);

            switch ( i ) {
               case 0: {
                  cell.setCellValue(new HSSFRichTextString("store")); // a
                  break;
               }

               case 1: {
                  cell.setCellValue(new HSSFRichTextString("websites")); //b
                  break;
               }

               case 2: {
                  cell.setCellValue(new HSSFRichTextString("attribute_set")); //c
                  break;
               }

               case 3: {
                  cell.setCellValue(new HSSFRichTextString("type")); //d
                  break;
               }

               case 4: {
                  cell.setCellValue(new HSSFRichTextString("category_ids")); //e
                  break;
               }

               case 5: {
                  cell.setCellValue(new HSSFRichTextString("sku")); //f
                  break;
               }

               case 6: {
                  cell.setCellValue(new HSSFRichTextString("has_options")); //g
                  break;
               }

               case 7: {
                  cell.setCellValue(new HSSFRichTextString("name")); //h
                  break;
               }

               case 8: {
                  cell.setCellValue(new HSSFRichTextString("meta_title")); //i
                  break;
               }

               case 9: {
                  cell.setCellValue(new HSSFRichTextString("meta_description"));//j
                  break;
               }

               case 10: {
                  cell.setCellValue(new HSSFRichTextString("image")); //k
                  break;
               }

               case 11: {
                  cell.setCellValue(new HSSFRichTextString("small_image")); //l
                  break;
               }

               case 12: {
                  cell.setCellValue(new HSSFRichTextString("thumbnail")); //m
                  break;
               }

               case 13: {
                  cell.setCellValue(new HSSFRichTextString("url_key")); //n
                  break;
               }

               case 14: {
                  cell.setCellValue(new HSSFRichTextString("url_path")); //o
                  break;
               }

               case 15: {
                  cell.setCellValue(new HSSFRichTextString("custom_design")); //p
                  break;
               }

               case 16: {
                  cell.setCellValue(new HSSFRichTextString("page_layout")); //Q
                  break;
               }

               case 17: {
                  cell.setCellValue(new HSSFRichTextString("options_container"));//r
                  break;
               }

               case 18: {
                  cell.setCellValue(new HSSFRichTextString("gift_message_available"));//s
                  break;
               }

               case 19: {
                  cell.setCellValue(new HSSFRichTextString("price"));//t
                  break;
               }

               case 20: {
                  cell.setCellValue(new HSSFRichTextString("special_price"));//u
                  break;
               }

               case 21: {
                  cell.setCellValue(new HSSFRichTextString("cost"));//v
                  break;
               }

               case 22: {
                  cell.setCellValue(new HSSFRichTextString("weight"));//w
                  break;
               }

               case 23: {
                  cell.setCellValue(new HSSFRichTextString("description"));//x
                  break;
               }

               case 24: {
                  cell.setCellValue(new HSSFRichTextString("short_description"));//y
                  break;
               }

               case 25: {
                  cell.setCellValue(new HSSFRichTextString("meta_keyword"));//z
                  break;
               }

               case 26: {
                  cell.setCellValue(new HSSFRichTextString("custom_layout_update"));//aa
                  break;
               }

               case 27: {
                  cell.setCellValue(new HSSFRichTextString("manufacturer"));//ab
                  break;
               }

               case 28: {
                  cell.setCellValue(new HSSFRichTextString("status"));//ac
                  break;
               }

               case 29: {
                  cell.setCellValue(new HSSFRichTextString("tax_class_id"));//ad
                  break;
               }

               case 30: {
                  cell.setCellValue(new HSSFRichTextString("visibility"));//ae
                  break;
               }

               case 31: {
                  cell.setCellValue(new HSSFRichTextString("enable_googlecheckout"));//af
                  break;
               }

               case 32: {
                  cell.setCellValue(new HSSFRichTextString("related_targetrule_position_limit"));//ag
                  break;
               }

               case 33: {
                  cell.setCellValue(new HSSFRichTextString("related_targetrule_position_behavior"));//ah
                  break;
               }

               case 34: {
                  cell.setCellValue(new HSSFRichTextString("upsell_targetrule_position_limit"));//ai
                  break;
               }

               case 35: {
                  cell.setCellValue(new HSSFRichTextString("upsell_targetrule_position_behavior"));//aj
                  break;
               }

               case 36: {
                  cell.setCellValue(new HSSFRichTextString("special_from_date"));//ak
                  break;
               }

               case 37: {
                  cell.setCellValue(new HSSFRichTextString("special_to_date"));//al
                  break;
               }

               case 38: {
                  cell.setCellValue(new HSSFRichTextString("news_from_date"));//am
                  break;
               }

               case 39: {
                  cell.setCellValue(new HSSFRichTextString("news_to_date"));//an
                  break;
               }

               case 40: {
                  cell.setCellValue(new HSSFRichTextString("custom_design_from"));//ao
                  break;
               }

               case 41: {
                  cell.setCellValue(new HSSFRichTextString("custom_design_to"));//ap
                  break;
               }

               case 42: {
                  cell.setCellValue(new HSSFRichTextString("qty"));//aq
                  break;
               }

               case 43: {
                  cell.setCellValue(new HSSFRichTextString("min_qty"));//ar
                  break;
               }

               case 44: {
                  cell.setCellValue(new HSSFRichTextString("use_config_min_qty"));//as
                  break;
               }

               case 45: {
                  cell.setCellValue(new HSSFRichTextString("is_qty_decimal"));//at
                  break;
               }

               case 46: {
                  cell.setCellValue(new HSSFRichTextString("backorders"));//au
                  break;
               }

               case 47: {
                  cell.setCellValue(new HSSFRichTextString("use_config_backorders"));//av
                  break;
               }

               case 48: {
                  cell.setCellValue(new HSSFRichTextString("min_sale_qty"));//aw
                  break;
               }

               case 49: {
                  cell.setCellValue(new HSSFRichTextString("use_config_min_sale_qty"));//ax
                  break;
               }

               case 50: {
                  cell.setCellValue(new HSSFRichTextString("max_sale_qty"));//ay
                  break;
               }

               case 51: {
                  cell.setCellValue(new HSSFRichTextString("use_config_max_sale_qty"));//az
                  break;
               }

               case 52: {
                  cell.setCellValue(new HSSFRichTextString("is_in_stock"));//ba
                  break;
               }

               case 53: {
                  cell.setCellValue(new HSSFRichTextString("low_stock_date"));//bb
                  break;
               }

               case 54: {
                  cell.setCellValue(new HSSFRichTextString("notify_stock_qty"));//bc
                  break;
               }

               case 55: {
                  cell.setCellValue(new HSSFRichTextString("use_config_notify_stock_qty"));//bd
                  break;
               }

               case 56: {
                  cell.setCellValue(new HSSFRichTextString("manage_stock"));//be
                  break;
               }

               case 57: {
                  cell.setCellValue(new HSSFRichTextString("use_config_manage_stock"));//bf
                  break;
               }

               case 58: {
                  cell.setCellValue(new HSSFRichTextString("stock_status_changed_automatically"));//bg
                  break;
               }

               case 59: {
                  cell.setCellValue(new HSSFRichTextString("product_name"));//bh
                  break;
               }

               case 60: {
                  cell.setCellValue(new HSSFRichTextString("store_id"));//bi
                  break;
               }

               case 61: {
                  cell.setCellValue(new HSSFRichTextString("product_type_id"));//bj
                  break;
               }

               case 62: {
                  cell.setCellValue(new HSSFRichTextString("product_status_changed"));//bk
                  break;
               }

               case 63: {
                  cell.setCellValue(new HSSFRichTextString("product_changed_websites"));//bl
                  break;
               }

               case 64: {
                  cell.setCellValue(new HSSFRichTextString("image_label"));//bm
                  break;
               }

               case 65: {
                  cell.setCellValue(new HSSFRichTextString("small_image_label"));//bn
                  break;
               }

               case 66: {
                  cell.setCellValue(new HSSFRichTextString("thumbnail_label"));//bo
                  break;
               }

               case 67: {
                  cell.setCellValue(new HSSFRichTextString("tank_size"));//bp
                  break;
               }

               case 68: {
                  cell.setCellValue(new HSSFRichTextString("compressor_type"));//bq
                  break;
               }

               case 69: {
                  cell.setCellValue(new HSSFRichTextString("mobility"));//br
                  break;
               }

               case 70: {
                  cell.setCellValue(new HSSFRichTextString("pump_lubrication"));//bs
                  break;
               }

               case 71: {
                  cell.setCellValue(new HSSFRichTextString("voltage"));//bt
                  break;
               }

               case 72: {
                  cell.setCellValue(new HSSFRichTextString("pump_type"));//bu
                  break;
               }

               case 73: {
                  cell.setCellValue(new HSSFRichTextString("horsepower"));//bv
                  break;
               }

               case 74: {
                  cell.setCellValue(new HSSFRichTextString("diameter"));//bw
                  break;
               }

               case 75: {
                  cell.setCellValue(new HSSFRichTextString("angle"));//bx
                  break;
               }

               case 76: {
                  cell.setCellValue(new HSSFRichTextString("cutting_length"));//by
                  break;
               }

               case 77: {
                  cell.setCellValue(new HSSFRichTextString("total_length"));//bz
                  break;
               }

               case 78: {
                  cell.setCellValue(new HSSFRichTextString("router_bits_shank"));//ca
                  break;
               }

               case 79: {
                  cell.setCellValue(new HSSFRichTextString("amperage"));//cb
                  break;
               }

               case 80: {
                  cell.setCellValue(new HSSFRichTextString("upc"));//cc
                  break;
               }

               case 81: {
                  cell.setCellValue(new HSSFRichTextString("tool_weight"));//cd
                  break;
               }

               case 82: {
                  cell.setCellValue(new HSSFRichTextString("stroke_length"));//ce
                  break;
               }

               case 83: {
                  cell.setCellValue(new HSSFRichTextString("no_load_speed"));//cf
                  break;
               }

               case 84: {
                  cell.setCellValue(new HSSFRichTextString("item_condition"));//cg
                  break;
               }

               case 85: {
                  cell.setCellValue(new HSSFRichTextString("keyless_blade_clamp"));//ch
                  break;
               }

               case 86: {
                  cell.setCellValue(new HSSFRichTextString("cord_type"));//ci
                  break;
               }

               case 87: {
                  cell.setCellValue(new HSSFRichTextString("keyless_adj_shoe"));//cj
                  break;
               }

               case 88: {
                  cell.setCellValue(new HSSFRichTextString("adjustable_handle"));//ck
                  break;
               }

               case 89: {
                  cell.setCellValue(new HSSFRichTextString("bevel_stops"));//cl
                  break;
               }

               case 90: {
                  cell.setCellValue(new HSSFRichTextString("miter_stops"));//cm
                  break;
               }

               case 91: {
                  cell.setCellValue(new HSSFRichTextString("blade_diameter"));//cn
                  break;
               }

               case 92: {
                  cell.setCellValue(new HSSFRichTextString("arbor"));//co
                  break;
               }
            }
         }
      }

      return ++rowNum;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
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
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    *
    * @return The new row of the spreadsheet.
    */
   private HSSFRow createRow(int rowNum)
   {
      HSSFRow row = null;

      if ( m_Sheet != null )
         row = m_Sheet.createRow(rowNum);

      return row;
   }

   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;

      if ( m_OraConn != null ) {
         try {
            sql.append("select\r\n");
            sql.append("   item.item_id, item.description, item.weight, upc_code, ");
            sql.append("   nvl(vendor_shortname.name, vendor.name) vnd_name, ");
            sql.append("   bmi_item.main_image || '_sm.gif' img_sm, ");
            sql.append("   bmi_item.main_image || '.gif' img_md, ");
            sql.append("   bmi_item.main_image || '_lg.gif' img_lg, ");
            sql.append("   cust_procs.GETSELLPRICE(?, item.item_id) cost, ");
            sql.append("   cust_procs.GETRETAILPRICE(?, item.item_id) retail ");
            sql.append("from ");
            sql.append("   customer ");
            sql.append("join cust_warehouse on cust_warehouse.customer_id = customer.customer_id ");
            sql.append("join item_warehouse on item_warehouse.warehouse_id = cust_warehouse.warehouse_id and item_warehouse.in_catalog = 1 ");
            sql.append("join item on item.item_id = item_warehouse.item_id and ");
            sql.append("   item.item_type_id = 1 and item.disp_id = 1 ");
            sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id and ship_unit.unit <> 'AST' ");
            sql.append("join bmi_item on bmi_item.item_id = item.item_id ");
            sql.append("join vendor on vendor.vendor_id = item.vendor_id ");
            sql.append("left outer join vendor_shortname on vendor_shortname.vendor_id = vendor.vendor_id ");
            sql.append("left outer join item_upc on item_upc.item_id = item.item_id and item_upc.primary_upc = 1 ");
            sql.append("where ");
            sql.append("   customer.customer_id = ? ");
            sql.append("order by item.item_id");

            m_ItemData = m_OraConn.prepareStatement(sql.toString());
            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("exception:", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("acetool.prepareStatements - null oracle connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();

      fileName.append("acetool-export");
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      HSSFCellStyle styleText = null;        // Text left justified
      HSSFCellStyle styleInt = null;         // Style with 0 decimals
      HSSFCellStyle styleDouble2 = null;     // numeric #,##0.00
      HSSFCellStyle styleDouble3 = null;     // numeric #,##0.000
      HSSFDataFormat format = null;

      format = m_Wrkbk.createDataFormat();

      styleText = m_Wrkbk.createCellStyle();
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleDouble2 = m_Wrkbk.createCellStyle();
      styleDouble2.setAlignment(HorizontalAlignment.RIGHT);
      styleDouble2.setDataFormat(format.getFormat("#,##0.00"));

      styleDouble3 = m_Wrkbk.createCellStyle();
      styleDouble3.setAlignment(HorizontalAlignment.RIGHT);
      styleDouble3.setDataFormat(format.getFormat("#,##0.000"));

      m_CellStyles = new HSSFCellStyle[] {
         styleText,
         styleInt,
         styleDouble2,
         styleDouble3
      };

      styleText = null;
      styleInt = null;
      styleDouble2 = null;
      styleDouble3 = null;
   }
}

