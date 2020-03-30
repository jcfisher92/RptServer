package com.emerywaterhouse.rpt.export;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptProcessor;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.utils.StringFormat;
import com.emerywaterhouse.websvc.Param;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import org.apache.log4j.Logger;

public class RetailCatalog extends Report
{
  private static final String WHERE_ITEM = "and item.item_id = '%s'\r\n";
  private static final String FLC_JOIN = "   and flc.flc_id = '%s' \r\n";
  private static final String MDC_JOIN = "   and mdc.mdc_id = '%s' \r\n";
  private static final String NRHA_JOIN = "   and nrha.nrha_id = '%s'\r\n";
  private static final String SKU_ELEMENT = "   <sku>%s</sku>\r\n";
  private static final String DESC_ELEMENT = "   <description>%s</description>\r\n";
  private static final String UPC_ELEMENT = "   <upc>%s</upc>\r\n";
  private static final String VNDSKU_ELEMENT = "   <vendorSku>%s</vendorSku>\r\n";
  private static final String VNDID_ELEMENT = "      <id>%d</id>\r\n";
  private static final String VNDNAME_ELEMENT = "      <name>%s</name>\r\n";
  private static final String IMGSM_ELEMENT = "   <imageUrlSm>%s</imageUrlSm>\r\n";
  private static final String IMGMD_ELEMENT = "   <imageUrlMd>%s</imageUrlMd>\r\n";
  private static final String IMGLG_ELEMENT = "   <imageUrlLg>%s</imageUrlLg>\r\n";
  private static final String BRKCASE_ELEMENT = "   <brokenCase>%s</brokenCase>\r\n";
  private static final String DLRPACK_ELEMENT = "   <dealerPack>%d</dealerPack>\r\n";
  private static final String PACKOF_ELEMENT = "   <packOf>%d</packOf>\r\n";
  private static final String COST_ELEMENT = "   <cost>%.3f</cost>\r\n";
  private static final String RETAIL_ELEMENT = "   <retail>%.2f</retail>\r\n";
  private static final String LENGTH_ELEMENT = "   <length>%.2f</length>\r\n";
  private static final String WIDTH_ELEMENT = "   <width>%.2f</width>\r\n";
  private static final String HEIGHT_ELEMENT = "   <height>%.2f</height>\r\n";
  private static final String WEIGHT_ELEMENT = "   <weight>%.2f</weight>\r\n";
  private static final String CUBE_ELEMENT = "   <cube>%.2f</cube>\r\n";
  private static final String UOM_ELEMENT = "   <uom>%s</uom>\r\n";
  private static final String FLC_ELEMENT = "   <flc>%s</flc>\r\n";
  private static final String MDC_ELEMENT = "   <mdc>%s</mdc>\r\n";
  private static final String DC_ELEMENT = "   <dc>%s</dc>\r\n";
  private static final String NRHA_ELEMENT = "   <nrha>%s</nrha>\r\n";
  private static final String BRAND_ELEMENT = "   <brandName>%s</brandName>\r\n";
  private static final String NOUN_ELEMENT = "   <noun>%s</noun>\r\n";
  private static final String MOD_ELEMENT = "   <modifier>%s</modifier>\r\n";
  private static final String RETA_ELEMENT = "   <retailA>%.2f</retailA>\r\n";
  private static final String RETB_ELEMENT = "   <retailB>%.2f</retailB>\r\n";
  private static final String RETC_ELEMENT = "   <retailC>%.2f</retailC>\r\n";
  private static final String RETD_ELEMENT = "   <retailD>%.2f</retailD>\r\n";
  private PreparedStatement m_AttrData;
  private PreparedStatement m_BulletData;
  private PreparedStatement m_CatData;
  private PreparedStatement m_HazardData;
  private PreparedStatement m_ItemData;
  private PreparedStatement m_PriceData;
  private PreparedStatement m_LocData;
  private PreparedStatement m_ParentLocData;
  private PreparedStatement m_CustWarehouse;
  private String m_CustId;
  private String m_PosVnd;
  private String m_DataFmt;
  private String m_DataSrc;
  private String m_SrcId;
  private int m_Dc;
  private boolean m_Overwrite;
  private String m_EscapeFormat;
  private boolean m_AllHazmat;
  private boolean m_IncludeAsst;
  private String m_DescrSrc;
  private boolean m_IncludeRetails;
  
  public RetailCatalog()
  {
    this.m_CustId = "";
    this.m_PosVnd = "";
    this.m_Dc = 1;
    this.m_Overwrite = false;
    this.m_EscapeFormat = "html";
    this.m_DataSrc = "";
    this.m_SrcId = "";
    this.m_DataFmt = "xml";
    this.m_AllHazmat = false;
    this.m_IncludeAsst = true;
    this.m_DescrSrc = "item";
  }
  
  public void finalize()
    throws Throwable
  {
    this.m_CustId = null;
    this.m_PosVnd = null;
    this.m_DataFmt = null;
    this.m_DataSrc = null;
    this.m_SrcId = null;
    
    super.finalize();
  }
  
  private boolean buildOutputFile()
    throws FileNotFoundException
  {
    FileOutputStream outFile = null;
    boolean result = false;
    
    outFile = new FileOutputStream(this.m_FilePath + (String)this.m_FileNames.get(0), false);
    try
    {
      if (this.m_DataFmt.equals("xml")) 
        result = buildXml(outFile);
    }
    catch (Exception ex)
    {
      this.m_ErrMsg.append("Your report had the following errors: \r\n");
      this.m_ErrMsg.append(ex.getClass().getName() + "\r\n");
      this.m_ErrMsg.append(ex.getMessage());
      
      log.fatal("exception:", ex);
    }
    finally
    {
      try {
        outFile.close();
      }
      catch (Exception e)
      {
        log.error(e);
      }
      outFile = null;
    }
    return result;
  }
  
  private boolean buildXml(FileOutputStream outFile)
    throws Exception
  {
    boolean result = false;
    ResultSet itemData = null;
    ResultSet rsPriceData = null;
    ResultSet catText = null;
    StringBuffer xml = new StringBuffer();
    String itemId = null;
    int itemEaId = -1;
    String desc = null;
    String upc = null;
    String brandName = null;
    String noun = null;
    String modifier = null;
    
    this.m_ItemData.setInt(1, this.m_Dc);
    itemData = this.m_ItemData.executeQuery();
    try
    {
      xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
      xml.append("<catalog xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
      xml.append("xmlns =\"http://www.emeryonline.com/catalog\" ");
      xml.append("xsi:schemaLocation=\"http://www.emeryonline.com/catalog ");
      xml.append("http://www.emeryonline.com/catalog/catalogexp.xsd\" >\r\n");
      outFile.write(xml.toString().getBytes());
      xml.setLength(0);
      
      while ((itemData.next()) && (this.m_Status == 1)) {
        itemId = itemData.getString("item_id");
        itemEaId = itemData.getInt("item_ea_id");
        
        // try to get pricing data, continue if you can't
        m_PriceData.setString(1, m_CustId);
        m_PriceData.setString(2, m_CustId);
        if (m_IncludeRetails) {
	        m_PriceData.setInt(3, m_Dc);
	        m_PriceData.setInt(4, m_Dc);
	        m_PriceData.setInt(5, m_Dc);
	        m_PriceData.setInt(6, m_Dc);
	        m_PriceData.setString(7, itemId);
	        m_PriceData.setString(8, m_CustId);
        } else {
	        m_PriceData.setString(3, itemId);
	        m_PriceData.setString(4, m_CustId);
        }
        try {
           rsPriceData = m_PriceData.executeQuery();
           if (!rsPriceData.next()) {
           	continue; // skip to the next item, we didn't get any pricing for this one.
           }
        } catch (Exception e) { // problem getting a price
      	  //log.error("[RetailCatalog#buildXml] Could not price item: " + itemId + " for customer: " + m_CustId + " for whs: " + m_Dc + " - " + e.getMessage()); // Spammy
      	  m_EdbConn.rollback();
        	  continue;
        }
        
        if (this.m_DescrSrc.equalsIgnoreCase("web")) 
          desc = itemData.getString("web_descr");
        else {
          desc = itemData.getString("description");
        }
        upc = itemData.getString("upc_code");
        brandName = itemData.getString("brand_name");
        noun = itemData.getString("noun");
        modifier = itemData.getString("modifier");
        
        setCurAction(String.format("processing item %s for customer %s", new Object[] { itemId, this.m_CustId }));
        if (upc == null) {
          upc = "";
        }
        if (brandName == null) {
          brandName = "";
        }
        desc = normalize(desc);
        brandName = normalize(brandName);
        noun = normalize(noun);
        modifier = normalize(modifier);
        
        xml.append("<catalogItem>\r\n");
        xml.append(String.format("   <sku>%s</sku>\r\n", new Object[] { itemId }));
        xml.append(String.format("   <description>%s</description>\r\n", new Object[] { desc }));
        xml.append(String.format("   <upc>%s</upc>\r\n", new Object[] { upc }));
        xml.append(String.format("   <vendorSku>%s</vendorSku>\r\n", new Object[] { normalize(itemData.getString("vendor_item_num")) }));
        xml.append("   <vendor>\r\n");
        xml.append(String.format("      <id>%d</id>\r\n", new Object[] { Integer.valueOf(itemData.getInt("vendor_id")) }));
        xml.append(String.format("      <name>%s</name>\r\n", new Object[] { normalize(itemData.getString("name")) }));
        xml.append("   </vendor>\r\n");
        xml.append(String.format("   <imageUrlSm>%s</imageUrlSm>\r\n", new Object[] { itemData.getString("img_url_sm") }));
        xml.append(String.format("   <imageUrlMd>%s</imageUrlMd>\r\n", new Object[] { itemData.getString("img_url_md") }));
        xml.append(String.format("   <imageUrlLg>%s</imageUrlLg>\r\n", new Object[] { itemData.getString("img_url_lg") }));
        xml.append(String.format("   <brokenCase>%s</brokenCase>\r\n", new Object[] { itemData.getString("broken_case") }));
        xml.append(String.format("   <dealerPack>%d</dealerPack>\r\n", new Object[] { Integer.valueOf(itemData.getInt("retail_pack")) }));
        xml.append(String.format("   <packOf>%d</packOf>\r\n", new Object[] { Integer.valueOf(itemData.getInt("packof")) }));
        xml.append(String.format("   <cost>%.3f</cost>\r\n", new Object[] { Double.valueOf(rsPriceData.getDouble("cost")) }));
        xml.append(String.format("   <retail>%.2f</retail>\r\n", new Object[] { Double.valueOf(rsPriceData.getDouble("retail")) }));
        xml.append(String.format("   <length>%.2f</length>\r\n", new Object[] { Double.valueOf(itemData.getDouble("length")) }));
        xml.append(String.format("   <width>%.2f</width>\r\n", new Object[] { Double.valueOf(itemData.getDouble("width")) }));
        xml.append(String.format("   <height>%.2f</height>\r\n", new Object[] { Double.valueOf(itemData.getDouble("height")) }));
        xml.append(String.format("   <weight>%.2f</weight>\r\n", new Object[] { Double.valueOf(itemData.getDouble("weight")) }));
        xml.append(String.format("   <cube>%.2f</cube>\r\n", new Object[] { Double.valueOf(itemData.getDouble("cube")) }));
        xml.append(String.format("   <uom>%s</uom>\r\n", new Object[] { itemData.getString("uom") }));
        xml.append(String.format("   <flc>%s</flc>\r\n", new Object[] { itemData.getString("flc_id") }));
        xml.append(String.format("   <mdc>%s</mdc>\r\n", new Object[] { itemData.getString("mdc_id") }));
        xml.append(String.format("   <dc>%s</dc>\r\n", new Object[] { Integer.valueOf(this.m_Dc) }));
        xml.append(String.format("   <nrha>%s</nrha>\r\n", new Object[] { itemData.getString("nrha_id") }));
        xml.append(String.format("   <brandName>%s</brandName>\r\n", new Object[] { brandName }));
        xml.append(String.format("   <noun>%s</noun>\r\n", new Object[] { noun }));
        xml.append(String.format("   <modifier>%s</modifier>\r\n", new Object[] { modifier }));
        xml.append(getAttributes(itemId));
        xml.append(getBullets(itemId));
        xml.append(getCategories(itemId));
        xml.append(getLocationData(itemEaId));
        xml.append(getHazardData(itemId));
        
        if (this.m_IncludeRetails) {
          xml.append(String.format("   <retailA>%.2f</retailA>\r\n", new Object[] { Double.valueOf(rsPriceData.getDouble("retaila")) }));
          xml.append(String.format("   <retailB>%.2f</retailB>\r\n", new Object[] { Double.valueOf(rsPriceData.getDouble("retailb")) }));
          xml.append(String.format("   <retailC>%.2f</retailC>\r\n", new Object[] { Double.valueOf(rsPriceData.getDouble("retailc")) }));
          xml.append(String.format("   <retailD>%.2f</retailD>\r\n", new Object[] { Double.valueOf(rsPriceData.getDouble("retaild")) }));
        }
        xml.append("</catalogItem>\r\n");
        
        outFile.write(xml.toString().getBytes());
        xml.setLength(0);
        closeRSet(catText);
      }
      xml.append("</catalog>");
      outFile.write(xml.toString().getBytes());
      result = true;
    }
    catch (SQLException ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      setCurAction(String.format("finished processing catalog data", new Object[0]));
      closeRSet(itemData);
      itemData = null;
      closeRSet(rsPriceData);
      rsPriceData = null;
    }
    return result;
  }
  
  private void closeStatements()
  {
    closeStmt(this.m_ItemData);
    closeStmt(this.m_PriceData);
    closeStmt(this.m_BulletData);
    closeStmt(this.m_AttrData);
    closeStmt(this.m_HazardData);
    closeStmt(this.m_CatData);
    closeStmt(this.m_LocData);
    closeStmt(this.m_ParentLocData);
    closeStmt(this.m_CustWarehouse);
    
    this.m_ItemData = null;
    this.m_PriceData = null;
    this.m_BulletData = null;
    this.m_AttrData = null;
    this.m_HazardData = null;
    this.m_CatData = null;
    this.m_LocData = null;
    this.m_ParentLocData = null;
    this.m_CustWarehouse = null;
  }
  
  public boolean createReport()
  {
    boolean created = false;
    this.m_Status = 1;
    try
    {
   	this.m_EdbConn = this.m_RptProc.getEdbConn();
      if (prepareStatements()) 
        created = buildOutputFile();
    }
    catch (Exception ex)
    {
      log.fatal("[RetailCatalog] exception:", ex);
    }
    finally
    {
      closeStatements();
      if (this.m_Status == 1) {
        this.m_Status = 2;
      }
    }
    return created;
  }
  
  private String getAttributes(String itemId)
    throws SQLException
  {
    StringBuffer xml = new StringBuffer();
    ResultSet rs = null;
    String uom = null;
    String name = null;
    String value = null;
    
    this.m_AttrData.setString(1, itemId);
    this.m_AttrData.setString(2, this.m_CustId);
    rs = this.m_AttrData.executeQuery();
    try
    {
      while (rs.next()) {
        name = normalize(rs.getString(1));
        value = normalize(rs.getString(2));
        uom = rs.getString(3);
        
        if (uom == null) 
          uom = "";
        else {
          uom = normalize(uom);
        }
        xml.append("   <attribute>\r\n");
        xml.append(String.format("      <name>%s</name>\r\n", new Object[] { name }));
        xml.append(String.format("      <value>%s</value>\r\n", new Object[] { value }));
        xml.append(String.format("      <uom>%s</uom>\r\n", new Object[] { uom }));
        xml.append("   </attribute>\r\n");
      }
    }
    catch (SQLException ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      DbUtils.closeDbConn(null, null, rs);
      rs = null;
      uom = null;
    }
    return xml.toString();
  }
  
  private String getBullets(String itemId)
    throws SQLException
  {
    StringBuffer xml = new StringBuffer();
    ResultSet rs = null;
    String bullet = null;
    
    this.m_BulletData.setString(1, itemId);
    this.m_BulletData.setString(2, this.m_CustId);
    rs = this.m_BulletData.executeQuery();
    try
    {
      while (rs.next()) {
        bullet = normalize(rs.getString(1));
        xml.append("   <bullet>\r\n");
        xml.append(String.format("      <point>%s</point>\r\n", new Object[] { bullet }));
        xml.append(String.format("      <seqNbr>%d</seqNbr>\r\n", new Object[] { Integer.valueOf(rs.getInt(2)) }));
        xml.append("   </bullet>\r\n");
      }
    }
    catch (SQLException ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      DbUtils.closeDbConn(null, null, rs);
      rs = null;
      bullet = null;
    }
    return xml.toString();
  }
  
  private String getCategories(String itemId)
    throws SQLException
  {
    StringBuffer xml = new StringBuffer();
    ResultSet rs = null;
    long tax2Id = 0L;
    long tax1Id = 0L;
    String taxDesc = null;
    
    this.m_CatData.setString(1, itemId);
    this.m_CatData.setString(2, this.m_CustId);
    rs = this.m_CatData.executeQuery();
    try
    {
      if (rs.next()) {
        tax2Id = rs.getLong("tax2Id");
        tax1Id = rs.getLong("tax1Id");
        taxDesc = normalize(rs.getString("tax1desc"));
        


        xml.append("   <category>\r\n");
        xml.append(String.format("      <id>%d</id>\r\n", new Object[] { Long.valueOf(tax1Id) }));
        xml.append(String.format("      <name>%s</name>\r\n", new Object[] { taxDesc }));
        xml.append("      <level>1</level>\r\n");
        xml.append("      <parentId></parentId>\r\n");
        xml.append("   </category>\r\n");
        


        taxDesc = normalize(rs.getString("tax2desc"));
        xml.append("   <category>\r\n");
        xml.append(String.format("      <id>%d</id>\r\n", new Object[] { Long.valueOf(tax2Id) }));
        xml.append(String.format("      <name>%s</name>\r\n", new Object[] { taxDesc }));
        xml.append("      <level>2</level>\r\n");
        xml.append(String.format("      <parentId>%d</parentId>\r\n", new Object[] { Long.valueOf(tax1Id) }));
        xml.append("   </category>\r\n");
        


        taxDesc = normalize(rs.getString("tax3desc"));
        xml.append("   <category>\r\n");
        xml.append(String.format("      <id>%d</id>\r\n", new Object[] { Long.valueOf(rs.getLong("tax3id")) }));
        xml.append(String.format("      <name>%s</name>\r\n", new Object[] { taxDesc }));
        xml.append("      <level>3</level>\r\n");
        xml.append(String.format("      <parentId>%d</parentId>\r\n", new Object[] { Long.valueOf(tax2Id) }));
        xml.append("   </category>\r\n");
      }
    }
    catch (SQLException ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      DbUtils.closeDbConn(null, null, rs);
      rs = null;
      taxDesc = null;
    }
    return xml.toString();
  }
  
  private String getHazardData(String itemId)
    throws SQLException
  {
    StringBuffer xml = new StringBuffer();
    ResultSet rs = null;
    String hazClass = null;
    String unnaCode = null;
    
    this.m_HazardData.setString(1, itemId);
    rs = this.m_HazardData.executeQuery();
    try
    {
      while (rs.next()) {
        hazClass = rs.getString(6);
        unnaCode = rs.getString(7);
        if (hazClass == null) {
          hazClass = "";
        }
        if (unnaCode == null) {
          unnaCode = "";
        }
        xml.append("   <hazardous>\r\n");
        xml.append(String.format("      <transport>%s</transport>\r\n", new Object[] { rs.getString(1) }));
        xml.append(String.format("      <aerosol>%s</aerosol>\r\n", new Object[] { rs.getString(2) }));
        xml.append(String.format("      <flammable>%s</flammable>\r\n", new Object[] { rs.getString(3) }));
        xml.append(String.format("      <flammablePlastic>%s</flammablePlastic>\r\n", new Object[] { rs.getString(4) }));
        xml.append(String.format("      <transCode>%s</transCode>\r\n", new Object[] { rs.getString(5) }));
        xml.append(String.format("      <class>%s</class>\r\n", new Object[] { hazClass }));
        xml.append(String.format("      <unnaCode>%s</unnaCode>\r\n", new Object[] { unnaCode }));
        xml.append("   </hazardous>\r\n");
      }
    }
    catch (SQLException ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      DbUtils.closeDbConn(null, null, rs);
      rs = null;
      unnaCode = null;
      hazClass = null;
    }
    return xml.toString();
  }
  
  private String getLocationData(int itemEaId)
    throws SQLException
  {
    StringBuffer xml = new StringBuffer();
    ResultSet rs = null;
    long parentLoc = 0L;
    
    this.m_LocData.setInt(1, itemEaId);
    rs = this.m_LocData.executeQuery();
    label327:
    try { if (rs.next()) {
        parentLoc = rs.getLong(3);
        xml.append("   <location>\r\n");
        xml.append(String.format("      <id>%d</id>\r\n", new Object[] { Long.valueOf(rs.getLong(1)) }));
        xml.append(String.format("      <name>%s</name>\r\n", new Object[] { normalize(rs.getString(2)) }));
        xml.append(String.format("      <parentId>%s</parentId>\r\n", new Object[] {
          parentLoc > 0L ? Long.toString(parentLoc) : "" }));
        xml.append("   </location>\r\n");
        
        if (parentLoc <= 0L) break label327; this.m_ParentLocData.setLong(1, parentLoc);
        rs = this.m_ParentLocData.executeQuery();
        
        if (rs.next()) {
          parentLoc = rs.getLong(3);
          xml.append("   <location>\r\n");
          xml.append(String.format("      <id>%d</id>\r\n", new Object[] { Long.valueOf(rs.getLong(1)) }));
          xml.append(String.format("      <name>%s</name>\r\n", new Object[] { normalize(rs.getString(2)) }));
          xml.append(String.format("      <parentId>%s</parentId>\r\n", new Object[] {
            parentLoc > 0L ? Long.toString(parentLoc) : "" }));
          xml.append("   </location>\r\n");
        }
      }
    }
    catch (SQLException ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      DbUtils.closeDbConn(null, null, rs);
      rs = null;
    }
    return xml.toString();
  }
  
  private String normalize(String str)
  {
    if (this.m_EscapeFormat.equals("xml")) {
      return normalizeForXml(str);
    }
    return StringFormat.normalize(str);
  }
  
  public static String normalizeForXml(String p_str)
  {
    StringBuffer strBuf = new StringBuffer();
    
    int nLen = p_str != null ? p_str.length() : 0;
    
    if (p_str != null) {
      for (int i = 0; i < nLen; i++) {
        char ch = p_str.charAt(i);
        
        switch (ch) {
        case '<': 
          strBuf.append("&lt;");
          break;
        case '>': 
          strBuf.append("&gt;");
          break;
        case '&': 
          strBuf.append("&amp;");
          break;
        case '"': 
          strBuf.append("&quot;");
          break;
        case '“': 
          strBuf.append("&quot;");
          break;
        case '”': 
          strBuf.append("&quot;");
          break;
        case '®': 
          strBuf.append('~');
          break;
        case '': 
          strBuf.append("...");
          break;
        case '': 
          strBuf.append("&amp;#145;");
          break;
        case '': 
          strBuf.append("&amp;#146;");
          break;
        case '': 
          strBuf.append("&quot;");
          break;
        case '': 
          strBuf.append("&quot;");
          break;
        case '•': 
          strBuf.append("&amp;#183;");
          break;
        case '': 
          strBuf.append("&amp;#149;");
          break;
        case '': 
          strBuf.append("-");
          break;
        case '': 
          strBuf.append("-");
          break;
        case '': 
          strBuf.append('`');
          break;
        case '¢': 
          strBuf.append("&amp;#162;");
          break;
        case '°': 
          strBuf.append("|");
          break;
        case '±': 
          strBuf.append("&amp;#177;");
          break;
        case '·': 
          strBuf.append("&amp;#183;");
          break;
        case 'º': 
          strBuf.append("|");
          break;
        case '¼': 
          strBuf.append("&amp;#188;");
          break;
        case '½': 
          strBuf.append("&amp;#189;");
          break;
        case '¾': 
          strBuf.append("&amp;#190;");
          break;
        case '¿': 
          strBuf.append('`');
          break;
        case 'Ø': 
          strBuf.append("&amp;#216;");
          break;
        case 'á': 
          strBuf.append("&amp;#225;");
          break;
        case 'é': 
          strBuf.append("&amp;#233;");
          break;
        case 'í': 
          strBuf.append("&amp;#237;");
          break;
        case 'ó': 
          strBuf.append("&amp;#243;");
          break;
        case '–': 
          strBuf.append("-");
          break;
        case '’': 
          strBuf.append("'");
          break;
        case '™': 
          strBuf.append('`');
          break;
        case '': 
          break;
        case 'â': 
          break;
        case ' ': 
          break;
        case 'Â': 
          break;
        case '\r': 
          break;
        case '\n': 
          break;
        default: 
          strBuf.append(ch);
        }
      }
    }
    return strBuf.toString();
  }
  
  private boolean prepareStatements()
  {
    StringBuffer sql = new StringBuffer(256);
    boolean isPrepared = false;
    
    if (this.m_EdbConn != null) 
    {
      try
      {
        sql.append("select \r\n");
        sql.append("   item_entity_attr.item_id, item_entity_attr.item_ea_id, item_entity_attr.description, upc_code, \r\n");
        sql.append("   viec.vendor_item_num, vendor.vendor_id, coalesce(vendor_shortname.name, vendor.name) as name, \r\n");
        //sql.append("   viec.vendor_item_num, vendor.vendor_id, nvl(vendor_shortname.name, vendor.name) as name, \r\n");
        sql.append("   web_item_ea.img_url_sm, \r\n");
        sql.append("   web_item_ea.img_url_md, \r\n");
        sql.append("   web_item_ea.img_url_lg, \r\n");
        sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 'yes', 'no') as broken_case, \r\n");
        sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 1, ejd_item_warehouse.stock_pack) as packof, \r\n");
        sql.append("   ejd_item_warehouse.Length, ejd_item_warehouse.Width, ejd_item_warehouse.Height, ejd_item.weight, \r\n");
        sql.append("   ejd_item_warehouse.Cube, retail_unit.unit uom, ejd_item.flc_id, mdc.mdc_id, mdc.nrha_id, \r\n");
        //sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 1, ejd_item_warehouse.stock_pack) as \"packof\", \r\n");
        //sql.append("   sku_master.\"Length\", sku_master.\"Width\", sku_master.\"Height\", ejd_item.weight, \r\n");
        //sql.append("   sku_master.\"Cube\", retail_unit.unit uom, ejd_item.flc_id, mdc.mdc_id, mdc.nrha_id, \r\n");
        sql.append("   item_entity_attr.retail_pack, web_item_ea.brand_name, noun, modifier, web_item_ea.web_descr \r\n");
        sql.append("from \r\n");
        sql.append("   item_entity_attr \r\n");
        sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
        sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and warehouse_id = ? and ejd_item_warehouse.in_catalog = 1 \r\n");
        sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id ");
        sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");
        sql.append("join web_item_ea on web_item_ea.item_ea_id = item_entity_attr.item_ea_id \r\n");
        sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id \r\n");
        sql.append("left outer join vendor_shortname on vendor_shortname.vendor_id = vendor.vendor_id ");
        sql.append("left outer join vendor_item_ea_cross viec on viec.item_ea_id = item_entity_attr.item_ea_id and \r\n");
        sql.append("   viec.vendor_id = item_entity_attr.vendor_id \r\n");
        sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and ejd_item_whs_upc.primary_upc = 1 and \r\n");
        sql.append("   ejd_item_whs_upc.warehouse_id = ejd_item_warehouse.warehouse_id ");
        //sql.append("left outer join sku_master on sku_master.\"SKU\" = item_entity_attr.item_id and sku_master.\"WAREHOUSE\" = warehouse.name \r\n");
        sql.append("join broken_case bc on bc.broken_case_id = ejd_item.broken_case_id \r\n");
        sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id \r\n");
        sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id \r\n ");
        if (!this.m_IncludeAsst) {
          sql.append(" and ship_unit.unit <> 'AST' ");
        }
        sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
        sql.append(this.m_DataSrc.equals("flc") ? String.format("   and flc.flc_id = '%s' \r\n", new Object[] { this.m_SrcId }) : "\r\n");
        sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
        sql.append(this.m_DataSrc.equals("mdc") ? String.format("   and mdc.mdc_id = '%s' \r\n", new Object[] { this.m_SrcId }) : "\r\n");
        sql.append("join nrha on nrha.nrha_id = mdc.nrha_id ");
        sql.append(this.m_DataSrc.equals("nrha") ? String.format("   and nrha.nrha_id = '%s'\r\n", new Object[] { this.m_SrcId }) : "\r\n");
        sql.append("where \r\n");
        sql.append("   item_entity_attr.item_type_id in (1)  ");
        sql.append("    and item_entity_attr.item_ea_id not in (  ");
        sql.append("      select item_entity_attr.item_ea_id from item_entity_attr ");
        sql.append("      join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id ");
        sql.append("      where ejd_item_warehouse.disp_id = 1 and ejd_item_warehouse.warehouse_id in (1, 2) ");
        sql.append("      and item_entity_attr.item_type_id = 8 ");
        sql.append("    ) ");
        sql.append(this.m_DataSrc.equals("item") ? String.format("and item_entity_attr.item_id = '%s'\r\n", new Object[] { this.m_SrcId }) : "\r\n");
        sql.append("order by item_entity_attr.item_id");
        this.m_ItemData = this.m_EdbConn.prepareStatement(sql.toString());
        
        sql.setLength(0);
        sql.append("select ");
        sql.append("   (select price from ejd_cust_procs.get_sell_price(?, item_entity_attr.item_ea_id)) as cost,  \r\n"); // cust_id
        sql.append("   ejd.cust_procs.getretailprice(?, item_entity_attr.item_id) retail \r\n"); // cust_id
        if ( m_IncludeRetails ) {
           sql.append(",  ejd.item_price_procs.todays_retaila(item_entity_attr.item_id, ?) as retaila, \r\n"); // dc
           sql.append("   ejd.item_price_procs.todays_retailb(item_entity_attr.item_id, ?) as retailb, \r\n"); // dc
           sql.append("   ejd.item_price_procs.todays_retailc(item_entity_attr.item_id, ?) as retailc, \r\n"); // dc
           sql.append("   ejd.item_price_procs.todays_retaild(item_entity_attr.item_id, ?) as retaild \r\n"); // dc
        }
        sql.append("from item_entity_attr where item_ea_id = (select code from ejd_item_procs.get_item_ea_id(?,?))  \r\n"); // item_id, cust_id
        this.m_PriceData = m_EdbConn.prepareStatement(sql.toString());

        sql.setLength(0);
        sql.append("select attr_name, attr_value, attr_uom ");
        sql.append("from web_item_ea_attribute ");
        sql.append("where item_ea_id = (select code from ejd_item_procs.get_item_ea_id(?,?)) "); // item_id, cust_id
        this.m_AttrData = this.m_EdbConn.prepareStatement(sql.toString());
        
        sql.setLength(0);
        sql.append("select trim(bullet_point), seq_nbr ");
        sql.append("from web_item_ea_bullet ");
        sql.append("where item_ea_id = (select code from ejd_item_procs.get_item_ea_id(?,?)) "); // item_id, cust_id
        sql.append("order by seq_nbr");
        this.m_BulletData = this.m_EdbConn.prepareStatement(sql.toString());
        
        sql.setLength(0);
        sql.append("select trans_method, aerosol, flammable, flammable_plastic, trans_code, ");
        sql.append("item_hazmat_view.class, unna_code ");
        sql.append("from item_hazmat_view ");
        sql.append("where item_id = ? ");
        sql.append(this.m_AllHazmat ? " " : " and excepted = 'N' ");
        sql.append("order by trans_method");
        this.m_HazardData = this.m_EdbConn.prepareStatement(sql.toString());
        
        sql.setLength(0);
        sql.append("select tax3.taxonomy_id as tax3id, tax3.taxonomy as tax3desc,");
        sql.append("tax2.taxonomy_id as tax2Id, tax2.taxonomy as tax2desc, ");
        sql.append("tax1.taxonomy_id as tax1Id, tax1.taxonomy as tax1desc ");
        sql.append("from item_entity_attr ");
        sql.append("join item_ea_taxonomy tax3 on tax3.taxonomy_id = item_entity_attr.taxonomy_id ");
        sql.append("left join item_ea_taxonomy tax2 on tax2.taxonomy_id = tax3.parent_id ");
        sql.append("left join item_ea_taxonomy tax1 on tax1.taxonomy_id = tax2.parent_id ");
        sql.append("where item_ea_id = (select code from ejd_item_procs.get_item_ea_id(?,?)) "); // item_id, cust_id
        this.m_CatData = this.m_EdbConn.prepareStatement(sql.toString());
        
        sql.setLength(0);
        sql.append("select ");
        sql.append("wie.location, cl.location_name, cl.parent_location ");
        sql.append("from web_item_ea wie ");
        sql.append("join catalog_location cl on cl.location = wie.location ");
        sql.append("where wie.item_ea_id = ?");
        this.m_LocData = this.m_EdbConn.prepareStatement(sql.toString());
        
        sql.setLength(0);
        sql.append("select warehouse_id from cust_warehouse where customer_id = ? ");
        sql.append("union ");
        sql.append("select ace_rsc_id ");
        sql.append("  from ace_rsc ");
        sql.append("  join ace_cust_xref on ace_cust_xref.ace_rsc = ace_rsc.sap_site_cd and customer_id = ?");
        this.m_CustWarehouse = this.m_EdbConn.prepareStatement(sql.toString());
        

        sql.setLength(0);
        sql.append("select ");
        sql.append("location, location_name, parent_location ");
        sql.append("from catalog_location ");
        sql.append("where location = ?");
        this.m_ParentLocData = this.m_EdbConn.prepareStatement(sql.toString());
        
        isPrepared = true;
      }
      catch (SQLException ex)
      {
        log.error("[RetailCatalog] exception:", ex);
      }
      finally
      {
        sql = null;
      }
    } 
    else {
      log.error("catalog.prepareStatements - null enterprisedb connection");
    }
    return isPrepared;
  }
  
  public void setParams(ArrayList<Param> params)
  {
    StringBuffer fileName = new StringBuffer();
    String tmp = Long.toString(System.currentTimeMillis());
    int pcount = params.size();
    Param param = null;
    File file = null;
    try
    {
      for (int i = 0; i < pcount; i++) {
        param = (Param)params.get(i);
        if (param.name.equals("dc")) {
          this.m_Dc = Integer.parseInt(param.value);
        }
        if (param.name.equals("cust")) {
          this.m_CustId = param.value;
        }
        if (param.name.equals("posvnd")) {
          this.m_PosVnd = param.value;
        }
        if (param.name.equals("datafmt")) {
          this.m_DataFmt = param.value;
        }
        if (param.name.equals("datasrc")) {
          this.m_DataSrc = param.value;
        }
        if (param.name.equals("srcid")) {
          this.m_SrcId = param.value;
        }
        if (param.name.equals("overwrite")) {
          this.m_Overwrite = (param.value.equalsIgnoreCase("true"));
        }
        if (param.name.equals("allhazmat")) {
          this.m_AllHazmat = (param.value.equalsIgnoreCase("true"));
        }
        if (param.name.equals("includeasst")) {
          this.m_IncludeAsst = (param.value.equalsIgnoreCase("true"));
        }
        if (param.name.equals("escapeformat")) {
          if (param.value.equalsIgnoreCase("xml"))
            this.m_EscapeFormat = "xml";
          else {
            this.m_EscapeFormat = "html";
          }
        }
        if (param.name.equals("descrsrc")) {
          this.m_DescrSrc = param.value;
        }
        if (param.name.equals("includeretails")) {
          this.m_IncludeRetails = (param.value.equalsIgnoreCase("true"));
        }
      }
      
      if (!this.m_Overwrite) {
        if ((this.m_PosVnd.equalsIgnoreCase("pacsoft")) || (this.m_PosVnd.equalsIgnoreCase("marginmaster")))
        {
          if (this.m_CustId.length() == 6) 
            fileName.append(String.format("%s-", new Object[] { this.m_CustId }));          
        }
        else {
          fileName.append(tmp);
          fileName.append("-");
          
          if ((this.m_DataSrc != null) && (this.m_DataSrc.trim().length() > 0)) {
            fileName.append(this.m_DataSrc.trim());
            fileName.append("-");
            
            if ((this.m_SrcId != null) && (this.m_SrcId.trim().length() > 0)) {
              fileName.append(this.m_SrcId);
              fileName.append("-");
            }
          }
        }
      }
      else
      {
        this.m_FilePath = String.format("%s%s/", new Object[] { this.m_FilePath, this.m_CustId });
        file = new File(this.m_FilePath);
        if (!file.exists()) {
          file.mkdir();
        }
      }
      fileName.append("emery-catalog.xml");
      this.m_FileNames.add(fileName.toString());
    }
    catch (Exception ex)
    {
      log.error("[RetailCatalog] exception:", ex);
    }
    finally
    {
      fileName = null;
      tmp = null;
      param = null;
      file = null;
    }
  }
  
  /*
  public static void main(String[] args) {
	  RetailCatalog cat = new RetailCatalog();

     Param p1 = new Param();
     p1.name = "cust";
     p1.value = "000001";
     Param p2 = new Param();
     p2.name = "dc";
     p2.value = "1";
     Param p3 = new Param();
     p3.name = "posvnd";
     p3.value = "marginmaster";
     Param p4 = new Param();
     p4.name = "datafmt";
     p4.value = "xml";
     // datasrc skipped
     // srcid skipped
     Param p5 = new Param();
     p5.name = "overwrite";
     p5.value = "false";
     Param p6 = new Param();
     p6.name = "allhazmat";
     p6.value = "true";
     Param p7 = new Param();
     p7.name = "includeasst";
     p7.value = "true";
     Param p8 = new Param();
     p8.name = "escapefmt";
     p8.value = "xml";
     ArrayList<Param> params = new ArrayList<Param>();
     params.add(p1);
     params.add(p2);
     params.add(p3);
     params.add(p4);
     params.add(p5);
     params.add(p6);
     params.add(p7);
     params.add(p8);
     
     cat.m_FilePath = "C:\\EXP\\";
     
  	java.util.Properties connProps = new java.util.Properties();
  	connProps.put("user", "ejd");
  	connProps.put("password", "boxer");
  	try {
  		cat.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);
  		cat.m_EdbConn.setAutoCommit(false);
  	} catch (Exception e) {
  		e.printStackTrace();
  	}
     
     cat.setParams(params);
     cat.createReport();
  }
  */
}
