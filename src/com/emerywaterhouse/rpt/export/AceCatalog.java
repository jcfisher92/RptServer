/**
 * File: Catalog.java
 * Description: Exports the catalog data in a variety of formats.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 11/07/2006
 * Last Update: $Id: Catalog.java,v 1.16 2014/08/06 16:37:03 sgillis Exp $
 *
 * History
 *    $Log: Catalog.java,v $
 *    Revision 1.16  2014/08/06 16:37:03  sgillis
 *    image url refactor
 *
 *    Revision 1.13  2012/10/05 13:36:58  jfisher
 *    Margin master updates
 *
 *    Revision 1.12  2012/07/11 17:30:56  jfisher
 *    in_catalog modification
 *
 *    Revision 1.11  2011/12/07 11:36:30  jfisher
 *    updated file naming routine
 *
 *    Revision 1.10  2010/07/15 20:36:34  prichter
 *    Added option to use the bmi_item.web_descr instead of item.description
 *
 *    Revision 1.9  2010/04/25 07:50:29  prichter
 *    More special character filters
 *
 *    Revision 1.8  2010/02/23 06:54:21  prichter
 *    Added more special character translations for the Retail Web site
 *
 *    Revision 1.7  2010/01/30 01:37:24  prichter
 *    Trim the spaces off bullet points
 *
 *    Revision 1.6  2010/01/25 01:05:33  prichter
 *    Retail web project.
 *
 *    Revision 1.5  2010/01/03 10:27:05  prichter
 *    Production versions
 *
 *    Revision 1.4  2009/06/29 18:36:26  npasnur
 *    Replaced catalog_item with bmi_item
 *
 *    Revision 1.3  2009/06/19 15:30:20  jfisher
 *    Revamped the class to export the new bmi data structures.  Removed some other format options.
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;


public class AceCatalog extends Report
{
   private static final String SKU_ELEMENT      = "   <sku>%s</sku>\r\n";
   private static final String DESC_ELEMENT     = "   <description>%s</description>\r\n";
   private static final String UPC_ELEMENT      = "   <upc>%s</upc>\r\n";
   private static final String VNDSKU_ELEMENT   = "   <vendorSku>%s</vendorSku>\r\n";
   private static final String VNDID_ELEMENT    = "      <id>%d</id>\r\n";
   private static final String VNDNAME_ELEMENT  = "      <name>%s</name>\r\n";
   private static final String IMGSM_ELEMENT    = "   <imageUrlSm>%s</imageUrlSm>\r\n";
   private static final String IMGMD_ELEMENT    = "   <imageUrlMd>%s</imageUrlMd>\r\n";
   private static final String IMGLG_ELEMENT    = "   <imageUrlLg>%s</imageUrlLg>\r\n";
   private static final String BRKCASE_ELEMENT  = "   <brokenCase>%s</brokenCase>\r\n";
   private static final String DLRPACK_ELEMENT  = "   <dealerPack>%d</dealerPack>\r\n";
   private static final String PACKOF_ELEMENT   = "   <packOf>%d</packOf>\r\n";
   private static final String COST_ELEMENT     = "   <cost>%.3f</cost>\r\n";
   private static final String RETAIL_ELEMENT   = "   <retail>%.2f</retail>\r\n";
   private static final String LENGTH_ELEMENT   = "   <length>%.2f</length>\r\n";
   private static final String WIDTH_ELEMENT    = "   <width>%.2f</width>\r\n";
   private static final String HEIGHT_ELEMENT   = "   <height>%.2f</height>\r\n";
   private static final String WEIGHT_ELEMENT   = "   <weight>%.2f</weight>\r\n";
   private static final String CUBE_ELEMENT     = "   <cube>%.2f</cube>\r\n";
   private static final String UOM_ELEMENT      = "   <uom>%s</uom>\r\n";
   private static final String FLC_ELEMENT      = "   <flc>%s</flc>\r\n";
   private static final String MDC_ELEMENT      = "   <mdc>%s</mdc>\r\n";
   private static final String NRHA_ELEMENT     = "   <nrha>%s</nrha>\r\n";
   private static final String BRAND_ELEMENT    = "   <brandName>%s</brandName>\r\n";
   private static final String NOUN_ELEMENT     = "   <noun>%s</noun>\r\n";
   private static final String MOD_ELEMENT      = "   <modifier>%s</modifier>\r\n";

   //
   // For margin master exported data
   private static final String RETA_ELEMENT     = "   <retailA>%.2f</retailA>\r\n";
   private static final String RETB_ELEMENT     = "   <retailB>%.2f</retailB>\r\n";
   private static final String RETC_ELEMENT     = "   <retailC>%.2f</retailC>\r\n";
   private static final String RETD_ELEMENT     = "   <retailD>%.2f</retailD>\r\n";

   private PreparedStatement m_AttrData;
   private PreparedStatement m_BulletData;
   private PreparedStatement m_CatData;
   private PreparedStatement m_HazardData;
   private PreparedStatement m_ItemData;
   private PreparedStatement m_LocData;
   private PreparedStatement m_ParentLocData;
   private PreparedStatement m_PriceData;

   //
   // Params
   private String m_CustId;          // The customer number for the report data to be run against.
   private String m_PosVnd;          // The POS vendor that requested the catalog.         
   private int m_Dc;                 // The customer's distribution center.
   private boolean m_Overwrite;      // Overwrite the file flag
   private boolean m_IncludeAsst;    // If true, items with a ship unit of AST will be included.  Default is true;
   private boolean m_IncludeRetails; // Include retails A - C

   /**
    *
    */
   public AceCatalog()
   {
      super();

      m_CustId = "";
      m_PosVnd = "";
      m_Dc = 11;
      m_Overwrite = false;      
      m_IncludeAsst = true;      
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_CustId = null;
      m_PosVnd = null;
      
      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    * @throws Exception 
    */
   private boolean buildOutputFile() throws Exception
   {
      FileOutputStream outFile = null;
      boolean result = false;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      result = buildXml(outFile);
      
      return result;
   }

   /**
    * Builds the catalog export in XML format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildXml(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      ResultSet rsItemData = null;
      ResultSet rsPriceData = null;
      StringBuffer xml = new StringBuffer();
      int itemEaId = 0;      
      String itemId = null;
      String desc = null;
      String upc = null;
      String brandName = null;
      String noun = null;
      String modifier = null;

      m_ItemData.setInt(1, m_Dc);
      rsItemData = m_ItemData.executeQuery();
      
      try {
         xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
         xml.append("<catalog xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
         xml.append("xmlns =\"http://www.emeryonline.com/catalog\" ");
         xml.append("xsi:schemaLocation=\"http://www.emeryonline.com/catalog ");
         xml.append("http://www.emeryonline.com/catalog/catalogexp.xsd\" >\r\n");
         outFile.write(xml.toString().getBytes());
         xml.setLength(0);

         while ( rsItemData.next() && m_Status == RptServer.RUNNING ) {
            itemEaId = rsItemData.getInt("item_ea_id");            
            itemId = rsItemData.getString("item_id");
           	desc = rsItemData.getString("web_descr");            
            upc = rsItemData.getString("upc_code");
            brandName = rsItemData.getString("brand_name");
            noun = rsItemData.getString("noun");
            modifier = rsItemData.getString("modifier");

            setCurAction(String.format("processing item %s for customer %s", itemId, m_CustId));

            m_PriceData.setString(1, m_CustId);
            m_PriceData.setString(2, m_CustId);
            m_PriceData.setInt(3, m_Dc);
            m_PriceData.setInt(4, itemEaId);
                        
            try {
	            rsPriceData = m_PriceData.executeQuery();
	            
	            if (!rsPriceData.next()) {
	            	continue; // skip to the next item, we didn't get any pricing for this one.
	            }
            } 
            
            catch ( Exception e ) { // problem getting a price            	
            	m_EdbConn.rollback();
            	continue;
            }

            //
            // Make sure there's no null text in the XML document.
            if ( upc == null )
               upc = "";

            if ( brandName == null )
               brandName = "";

            //
            // Normalize any strings that might have reserved characters in it.
            desc = normalize(desc);
            brandName = normalize(brandName);
            noun = normalize(noun);
            modifier = normalize(modifier);

            xml.append("<catalogItem>\r\n");
            xml.append(String.format(SKU_ELEMENT, itemId));
           	xml.append(String.format(DESC_ELEMENT, desc));
            xml.append(String.format(UPC_ELEMENT, upc));
            xml.append(String.format(VNDSKU_ELEMENT, normalize(rsItemData.getString("vendor_item_num"))));
            xml.append("   <vendor>\r\n");
            xml.append(String.format(VNDID_ELEMENT, rsItemData.getInt("vendor_id")));
            xml.append(String.format(VNDNAME_ELEMENT, normalize(rsItemData.getString("name"))));
            xml.append("   </vendor>\r\n");
            xml.append(String.format(IMGSM_ELEMENT, rsItemData.getString("img_url_sm")));
            xml.append(String.format(IMGMD_ELEMENT, rsItemData.getString("img_url_md")));
            xml.append(String.format(IMGLG_ELEMENT, rsItemData.getString("img_url_ld")));
            xml.append(String.format(BRKCASE_ELEMENT, rsItemData.getString("broken_case")));
            xml.append(String.format(DLRPACK_ELEMENT, rsItemData.getInt("retail_pack")));
            xml.append(String.format(PACKOF_ELEMENT, rsItemData.getInt("packof")));
            xml.append(String.format(COST_ELEMENT, rsPriceData.getDouble("cost")));
            xml.append(String.format(RETAIL_ELEMENT, rsPriceData.getDouble("retail")));
            xml.append(String.format(LENGTH_ELEMENT, rsItemData.getDouble("length")));
            xml.append(String.format(WIDTH_ELEMENT, rsItemData.getDouble("width")));
            xml.append(String.format(HEIGHT_ELEMENT, rsItemData.getDouble("height")));
            xml.append(String.format(WEIGHT_ELEMENT, rsItemData.getDouble("weight")));
            xml.append(String.format(CUBE_ELEMENT, rsItemData.getDouble("cube")));
            xml.append(String.format(UOM_ELEMENT, rsItemData.getString("uom")));
            xml.append(String.format(FLC_ELEMENT, rsItemData.getString("flc_id")));
            xml.append(String.format(MDC_ELEMENT, rsItemData.getString("mdc_id")));
            xml.append(String.format(NRHA_ELEMENT, rsItemData.getString("nrha_id")));
            xml.append(String.format(BRAND_ELEMENT, brandName));
            xml.append(String.format(NOUN_ELEMENT, noun));
            xml.append(String.format(MOD_ELEMENT, modifier));
            xml.append(getAttributes(itemEaId));
            xml.append(getBullets(itemEaId));
            xml.append(getCategories(itemEaId));
            xml.append(getLocationData(itemEaId));
            xml.append(getHazardData(itemId));

            if ( m_IncludeRetails ) {
               xml.append(String.format(RETA_ELEMENT, rsPriceData.getDouble("retail_a")));
               xml.append(String.format(RETB_ELEMENT, rsPriceData.getDouble("retail_b")));
               xml.append(String.format(RETC_ELEMENT, rsPriceData.getDouble("retail_c")));
               xml.append(String.format(RETD_ELEMENT, rsPriceData.getDouble("retail_d")));
            }

            xml.append("</catalogItem>\r\n");

            outFile.write(xml.toString().getBytes());
            xml.setLength(0);
         }

         xml.append("</catalog>");
         outFile.write(xml.toString().getBytes());
         result = true;
      }
      
      catch ( Exception ex ) {
         log.error("[AceCatalog]", ex);
      }

      finally {
         setCurAction(String.format("finished processing catalog data"));
         closeRSet(rsItemData);
         closeRSet(rsPriceData);
         rsItemData = null;
         rsPriceData = null;
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemData);
      closeStmt(m_BulletData);
      closeStmt(m_AttrData);
      closeStmt(m_HazardData);
      closeStmt(m_CatData);
      closeStmt(m_LocData);
      closeStmt(m_ParentLocData);
      closeStmt(m_PriceData);

      m_ItemData = null;
      m_BulletData = null;
      m_AttrData = null;
      m_HazardData = null;
      m_CatData = null;
      m_LocData = null;
      m_ParentLocData = null;
      m_PriceData = null;
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
      	m_EdbConn = m_RptProc.getEdbConn();
         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[Catalog]#createReport", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Creates the item attribute elements for specified item.
    *
    * @param itemId The item id to get the data for.
    * @return A string that contains the attribute XML.
    * @throws SQLException
    */
   private String getAttributes(int itemEaId) throws SQLException
   {
      StringBuffer xml = new StringBuffer();
      ResultSet rs = null;
      String uom = null;
      String name = null;
      String value = null;

      m_AttrData.setInt(1, itemEaId);      
      rs = m_AttrData.executeQuery();

      try {
         while ( rs.next() ) {
            name = normalize(rs.getString(1));
            value = normalize(rs.getString(2));
            uom = rs.getString(3);

            if ( uom == null )
               uom = "";
            else
               uom = normalize(uom);

            xml.append("   <attribute>\r\n");
            xml.append(String.format("      <name>%s</name>\r\n", name));
            xml.append(String.format("      <value>%s</value>\r\n", value));
            xml.append(String.format("      <uom>%s</uom>\r\n", uom));
            xml.append("   </attribute>\r\n");
         }
      }
      
      catch ( Exception ex ) {
         log.error("[AceCatalog]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         uom = null;
      }

      return xml.toString();
   }

   /**
    * Creates the item bullet elements for specified item.
    *
    * @param itemId The item id to get the data for.
    * @return A string that contains the bullet XML.
    * @throws SQLException
    */
   private String getBullets(int itemEaId) throws SQLException
   {
      StringBuffer xml = new StringBuffer();
      ResultSet rs = null;
      String bullet = null;

      m_BulletData.setInt(1, itemEaId);
      rs = m_BulletData.executeQuery();

      try {
         while ( rs.next() ) {
            bullet = normalize(rs.getString(1));
            xml.append("   <bullet>\r\n");
            xml.append(String.format("      <point>%s</point>\r\n", bullet));
            xml.append(String.format("      <seqNbr>%d</seqNbr>\r\n", rs.getInt(2)));
            xml.append("   </bullet>\r\n");
         }
      }
      
      catch ( Exception ex ) {
         log.error("[AceCatalog]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         bullet = null;
      }

      return xml.toString();
   }

   /**
    * Creates the category elements for specified item.
    *
    * @param itemId The item id to get the data for.
    * @return A string that contains the complete listing of the category element XML.
    *
    * @throws SQLException
    */
   private String getCategories(int itemEaId) throws SQLException
   {
      StringBuffer xml = new StringBuffer();
      ResultSet rs = null;
      long tax2Id = 0;
      long tax1Id = 0;
      String taxDesc = null;

      m_CatData.setInt(1, itemEaId);
      rs = m_CatData.executeQuery();

      //
      // Because there are always three taxonomy levels and only three taxonomy levels we
      // can write the query to pull the data in one shot with no iteration.
      try {
         if ( rs.next() ) {
            tax2Id = rs.getLong("tax2Id");
            tax1Id = rs.getLong("tax1Id");
            taxDesc = normalize(rs.getString("tax1desc"));

            //
            // Taxonomy level one
            xml.append("   <category>\r\n");
            xml.append(String.format("      <id>%d</id>\r\n", tax1Id));
            xml.append(String.format("      <name>%s</name>\r\n", taxDesc));
            xml.append("      <level>1</level>\r\n");
            xml.append("      <parentId></parentId>\r\n");
            xml.append("   </category>\r\n");

            //
            // Taxonomy level two
            taxDesc = normalize(rs.getString("tax2desc"));
            xml.append("   <category>\r\n");
            xml.append(String.format("      <id>%d</id>\r\n", tax2Id));
            xml.append(String.format("      <name>%s</name>\r\n", taxDesc));
            xml.append("      <level>2</level>\r\n");
            xml.append(String.format("      <parentId>%d</parentId>\r\n", tax1Id));
            xml.append("   </category>\r\n");

            //
            // Taxonomy level three
            taxDesc = normalize(rs.getString("tax3desc"));
            xml.append("   <category>\r\n");
            xml.append(String.format("      <id>%d</id>\r\n", rs.getLong("tax3id")));
            xml.append(String.format("      <name>%s</name>\r\n", taxDesc));
            xml.append("      <level>3</level>\r\n");
            xml.append(String.format("      <parentId>%d</parentId>\r\n", tax2Id));
            xml.append("   </category>\r\n");
         }
      }
      
      catch ( Exception ex ) {
         log.error("[AceCatalog]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         taxDesc = null;
      }

      return xml.toString();
   }

   /**
    * Creates the hazard data elements for specified item.
    *
    * @param itemId The item id to get the data for.
    * @return A string that contains the hazard data XML.
    * @throws SQLException
    */
   private String getHazardData(String itemId) throws SQLException
   {
      StringBuffer xml = new StringBuffer();
      ResultSet rs = null;
      String hazClass = null;
      String unnaCode = null;

      m_HazardData.setString(1, itemId);
      rs = m_HazardData.executeQuery();

      try {
         while ( rs.next() ) {
            hazClass = rs.getString(6);
            unnaCode = rs.getString(7);

            if ( hazClass == null )
               hazClass = "";

            if ( unnaCode == null )
               unnaCode = "";

            xml.append("   <hazardous>\r\n");
            xml.append(String.format("      <transport>%s</transport>\r\n", rs.getString(1)));
            xml.append(String.format("      <aerosol>%s</aerosol>\r\n", rs.getString(2)));
            xml.append(String.format("      <flammable>%s</flammable>\r\n", rs.getString(3)));
            xml.append(String.format("      <flammablePlastic>%s</flammablePlastic>\r\n", rs.getString(4)));
            xml.append(String.format("      <transCode>%s</transCode>\r\n", rs.getString(5)));
            xml.append(String.format("      <class>%s</class>\r\n", hazClass));
            xml.append(String.format("      <unnaCode>%s</unnaCode>\r\n", unnaCode));
            xml.append("   </hazardous>\r\n");
         }
      }
      
      catch ( Exception ex ) {
         log.error("[AceCatalog]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         unnaCode = null;
         hazClass = null;
      }

      return xml.toString();
   }

   /**
    * Outputs the location data for the file.
    *
    * @param itemId The item to get the data for.
    * @return The XML elements for the item.
    * @throws SQLException
    */
   private String getLocationData(int itemEaId) throws SQLException
   {
      StringBuffer xml = new StringBuffer();
      ResultSet rs = null;
      long parentLoc = 0;

      m_LocData.setInt(1, itemEaId);
      rs = m_LocData.executeQuery();

      try {
         if ( rs.next() ) {
            parentLoc = rs.getLong(3);
            xml.append("   <location>\r\n");
            xml.append(String.format("      <id>%d</id>\r\n", rs.getLong(1)));
            xml.append(String.format("      <name>%s</name>\r\n", normalize(rs.getString(2))));
            xml.append(String.format("      <parentId>%s</parentId>\r\n",
                  parentLoc > 0 ? Long.toString(parentLoc) : ""));
            xml.append("   </location>\r\n");

            if ( parentLoc > 0 ) {
               m_ParentLocData.setLong(1, parentLoc);
               rs = m_ParentLocData.executeQuery();

               //
               // Add the parent location data if it exists.  We only go up one location.
               if ( rs.next() ) {
                  parentLoc = rs.getLong(3);
                  xml.append("   <location>\r\n");
                  xml.append(String.format("      <id>%d</id>\r\n", rs.getLong(1)));
                  xml.append(String.format("      <name>%s</name>\r\n", normalize(rs.getString(2))));
                  xml.append(String.format("      <parentId>%s</parentId>\r\n",
                        parentLoc > 0 ? Long.toString(parentLoc) : ""));
                  xml.append("   </location>\r\n");
               }
            }
         }
      }
      
      catch ( Exception ex ) {
         log.error("[AceCatalog] ", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return xml.toString();
   }

   /**
   * This public method normalizes the given string to XML standards.  The
   * resulting xml can be parsed by xpp.
   *
   * @param p_str String - the string to normalize.
   * @return String - the normalized string.
   */
   public static String normalize(String p_str)
   {
      StringBuffer strBuf = new StringBuffer();
      char ch;


      int nLen = (p_str != null) ? p_str.length() : 0;

      if ( p_str != null ) {
         for ( int i = 0; i < nLen; i++ ) {
            ch = p_str.charAt(i);

            switch ( ch ) {
                case '<': {
                    strBuf.append("&lt;");
                    break;
                }
                case '>': {
                    strBuf.append("&gt;");
                    break;
                }
                case '&': {
                    strBuf.append("&amp;");
                    break;
                }
                case '"': {
                    strBuf.append("&quot;");
                    break;
                }
                case '“': { // open quote
                   strBuf.append("&quot;");
                   break;
                }
                case '”': { // close quote
                   strBuf.append("&quot;");
                   break;
                }
                case '®': {
                   strBuf.append('~');
                   break;
                }
                case 133 : { // elipse ...
               	 strBuf.append("...");
               	 break;
                }
                case 145 : { // yet another apostrophe
               	 strBuf.append("&amp;#145;");
               	 break;
                }
                case 146 : { //  apostrophe
               	 strBuf.append("&amp;#146;");
               	 break;
                }
                case 147 : {  // open double quote
               	 strBuf.append("&quot;");
               	 break;
                }
                case 148 : {  // close double quote
               	 strBuf.append("&quot;");
               	 break;
                }
                case '�' : {
               	 strBuf.append("&amp;#183;"); // unicode bullet point
               	 break;
                }
                case 149 : { // bullet point
               	 strBuf.append("&amp;#149;");
               	 break;
                }
                case 150 : {  // another long hyphen
               	 strBuf.append("-");
               	 break;
                }
                case 151 : {  // another long hyphen
               	 strBuf.append("-");
               	 break;
                }
                case 153 : { // ascii trademark
               	 strBuf.append('`');
               	 break;
                }
                case 162 : { // cent
               	 strBuf.append("&amp;#162;");
               	 break;
                }
                case 176: {   // degree U+00B0 (176)
                   strBuf.append("|");
                   break;
                }
                case 177 : { // plus or minus
               	 strBuf.append("&amp;#177;");
               	 break;
                }
                case 183 : {
               	 strBuf.append("&amp;#183;"); // bullet point
               	 break;
                }
                case 186: {   // degree ascii
                   strBuf.append("|");
                   break;
                }
                case 188 : { // 1/4 symbol
               	 strBuf.append("&amp;#188;");
               	 break;
                }
                case 189 : { // 1/2 symbol
               	 strBuf.append("&amp;#189;");
               	 break;
                }
                case 190 : { // 3/4 symbol
               	 strBuf.append("&amp;#190;");
               	 break;
                }
                case 191 : { // Upside down question mark - trademark?
               	 strBuf.append('`');
               	 break;
                }
                case 216 : { // theta
               	 strBuf.append("&amp;#216;");
               	 break;
                }
                case 225 : {	// lower a with accent
               	 strBuf.append("&amp;#225;");
               	 break;
                }
                case 233 : {	// lower e with accent - as in decor
               	 strBuf.append("&amp;#233;");
               	 break;
                }
                case 237 : {	// lower i with accent
               	 strBuf.append("&amp;#237;");
               	 break;
                }
                case 243 : {	// lower o with accent
               	 strBuf.append("&amp;#243;");
               	 break;
                }
                case 8211 : {   //  long hyphen
               	 //strBuf.append("&amp;#8211;");
               	 strBuf.append("-");
                   break;
                }
                case 8217 : { // another apostrophe
               	 strBuf.append("'");
               	 break;
                }
                case 8482 : {   // trademark U+2122 (8482)
                   strBuf.append('`');
                   break;
                }

                // just skip these next characters
                case 132:  // 2 commas
               	 break;
                case 226:  // 'a' with a hat
               	 break;
                case 160:  // mystery character
               	 break;
                case 194:  // 'A' with a hat
               	 break;

                case '\r':
                    break;
                case '\n':
                    break;

                default: { // default append char
               	 //System.out.println(ch + ": " + (int)ch);
                   strBuf.append(ch);
                }
            }
         }
      }

      return strBuf.toString();
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
            sql.append("select\r\n");
            sql.append("   item_entity_attr.item_ea_id, item_entity_attr.ejd_item_id, ");
            sql.append("   item_entity_attr.item_id, item_entity_attr.description, upc_code, viec.vendor_item_num, \r\n");
            sql.append("   vendor.vendor_id, nvl(vendor_shortname.name, vendor.name) as name, \r\n");
            sql.append("   'http://www.emeryonline.com/shared/images/catalog/small/' || web_item_ea.main_image || '_sm.gif' as img_url_sm, \r\n");
            sql.append("   'http://www.emeryonline.com/shared/images/catalog/medium/' || web_item_ea.main_image || '.gif' as img_url_md, \r\n");
            sql.append("   'http://www.emeryonline.com/shared/images/catalog/large/' || web_item_ea.main_image || '_lg.gif' as img_url_ld, \r\n");
            sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 'yes', 'no') as broken_case, \r\n");
            sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 1, ejd_item_warehouse.stock_pack) as \"packof\", \r\n");
            sql.append("   length, width, height, weight, cube, \r\n");
            sql.append("   retail_unit.unit uom, ejd_item.flc_id, mdc.mdc_id, mdc.nrha_id, \r\n");
            sql.append("   item_entity_attr.retail_pack, web_item_ea.brand_name, noun, modifier, web_item_ea.web_descr \r\n");
            sql.append("from \r\n");
            sql.append("   item_entity_attr \r\n");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and warehouse_id = ? and ejd_item_warehouse.in_catalog = 1 \r\n");
            sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id \r\n");
            sql.append("join web_item_ea on web_item_ea.item_ea_id = item_entity_attr.item_ea_id \r\n");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id \r\n");
            sql.append("left outer join vendor_shortname on vendor_shortname.vendor_id = vendor.vendor_id \r\n");
            sql.append("left outer join vendor_item_ea_cross viec on viec.item_ea_id = item_entity_attr.item_ea_id and \r\n");
            sql.append("   viec.vendor_id = item_entity_attr.vendor_id \r\n");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and ejd_item_whs_upc.primary_upc = 1 and \r\n");
            sql.append("   ejd_item_whs_upc.warehouse_id = ejd_item_warehouse.warehouse_id \r\n");            
            sql.append("join broken_case bc on bc.broken_case_id = ejd_item.broken_case_id \r\n");
            sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id \r\n");
            sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id \r\n ");

            if ( !m_IncludeAsst )
            	sql.append(" and ship_unit.unit <> 'AST' ");

            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");            
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join nrha on nrha.nrha_id = mdc.nrha_id ");
            sql.append("where item_entity_attr.item_type_id = 8 ");            
            sql.append("order by item_entity_attr.item_id, item_entity_attr.item_type_id ");
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
                        
            sql.setLength(0);
            sql.append("select ");
            sql.append("   (select price from ejd_cust_procs.get_sell_price(?, item_ea_id)) as cost, \r\n");
            sql.append("   ejd_price_procs.get_retail_price(?, item_ea_id) as retail, \r\n");
            sql.append("   retail_a, retail_b, retail_c, retail_d ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id and warehouse_id = ? ");
            sql.append("where item_ea_id = ?");
            m_PriceData = m_EdbConn.prepareStatement(sql.toString());

            //
            // Item attributes
            sql.setLength(0);
            sql.append("select attr_name, attr_value, attr_uom ");
            sql.append("from web_item_ea_attribute ");
            sql.append("where item_ea_id = ? ");
            m_AttrData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select trim(bullet_point), seq_nbr ");
            sql.append("from web_item_ea_bullet ");
            sql.append("where item_ea_id = ? ");
            sql.append("order by seq_nbr");
            m_BulletData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select trans_method, aerosol, flammable, flammable_plastic, trans_code, ");
            sql.append("item_hazmat_view.class, unna_code ");
            sql.append("from item_hazmat_view ");
            sql.append("where item_id = ? ");            
            sql.append("order by trans_method");
            m_HazardData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select tax3.taxonomy_id as tax3id, tax3.taxonomy as tax3desc,");
            sql.append("tax2.taxonomy_id as tax2Id, tax2.taxonomy as tax2desc, ");
            sql.append("tax1.taxonomy_id as tax1Id, tax1.taxonomy as tax1desc ");
            sql.append("from item_entity_attr ");
            sql.append("join item_ea_taxonomy tax3 on tax3.taxonomy_id = item_entity_attr.taxonomy_id ");
            sql.append("left join item_ea_taxonomy tax2 on tax2.taxonomy_id = tax3.parent_id ");
            sql.append("left join item_ea_taxonomy tax1 on tax1.taxonomy_id = tax2.parent_id ");
            sql.append("where item_ea_id = ? ");
            m_CatData = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select wie.location, cl.location_name, cl.parent_location ");
            sql.append("from web_item_ea wie ");
            sql.append("join catalog_location cl on cl.location = wie.location ");
            sql.append("where item_ea_id = ? ");
            m_LocData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select ");
            sql.append("location, location_name, parent_location ");
            sql.append("from catalog_location ");
            sql.append("where location = ?");
            m_ParentLocData = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( Exception ex ) {
            log.error("[Catalog]#prepareStatements", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[Catalog]#prepareStatements - null enterprisedb connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      String tmp = Long.toString(System.currentTimeMillis());
      int pcount = params.size();
      Param param = null;
      File file = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("dc") )
            m_Dc = Integer.parseInt(param.value);

         if ( param.name.equals("cust") )
            m_CustId = param.value;

         if ( param.name.equals("posvnd") )
            m_PosVnd = param.value;

         if ( param.name.equals("overwrite") )
            m_Overwrite = param.value.equalsIgnoreCase("true") ? true : false;

         if ( param.name.equals("includeasst") )
         	m_IncludeAsst = param.value.equalsIgnoreCase("true") ? true : false;

         if ( param.name.equals("includeretails") )
            m_IncludeRetails = param.value.equalsIgnoreCase("true") ? true : false;
      }

      //
      // Some customers want the same file name each time.  If that's the case, we
      // need to overwrite what we have.  For POS Vendors, we will overwrite, but with the name of the specific
      // customer.  That allows different files that still overwrite the previous customer specific file.
      if ( !m_Overwrite ) {
         if ( m_PosVnd.equalsIgnoreCase("pacsoft") || m_PosVnd.equalsIgnoreCase("marginmaster") ||  m_PosVnd.equalsIgnoreCase("aubuchon") ) {
            if ( m_CustId.length() == 6 )
               fileName.append(String.format("%s-", m_CustId));
         }
         else {
            fileName.append(tmp);
            fileName.append("-");            
         }
      }
      else {
         m_FilePath = String.format("%s%s/", m_FilePath, m_CustId);
         file = new File(m_FilePath);

         if ( !file.exists() )
            file.mkdir();
      }

      fileName.append("emery-catalog.xml");
      m_FileNames.add(fileName.toString());
   }         
}
