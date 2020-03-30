/**
 * File:
 * Description:
 *
 * @author
 *
 * Create Date:
 * Last Update:
 *
 * History:
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.ojb.broker.util.GUID;

import com.emerywaterhouse.oag.OagConst;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class ParData extends Report
{
   private static final int fmtBod = 0;
   private static final int fmtRet = 1;
      
   private String m_CustId;
   private int m_Dc;
   private byte m_Env;
   private int m_Format;
   private String m_VndId;
   
   private PreparedStatement m_ParData;
   
   /**
    * 
    */
   public ParData()
   {
      super();
      
      m_CustId = "";
      m_Dc = -1;
      m_Env = OagConst.bTEST;
      m_Format = -1;
      m_VndId = "";
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {
      m_CustId = null;
      m_VndId = null;
      
      super.finalize();
   }
     
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {      
      FileOutputStream outFile = null;      
      boolean result = false;
           
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         m_ParData.setString(1, m_CustId);
         m_ParData.setString(2, m_CustId);
         
         if ( m_Dc > 0 )
            m_ParData.setInt(3, m_Dc);
         
         //
         // Output the data in the specific format.  The ouput procs will
         // execute the query.
         switch ( m_Format ) {
            case fmtBod: { 
               outFile.write(parAsBod().getBytes());
               break;
            }
            
            case fmtRet: {
               outFile.write(parAsRet().getBytes());
               break;
            }
            
            default: {
               throw new Exception("unkown data format.");
            }
         }
         
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

         outFile = null;
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ParData);
            
      m_ParData = null;      
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
    * Formats the input java.util.Date to return a String date value in the Oagis DateTime format
    *  i.e. <xs:pattern value="\d\d\d\d-\d\d-\d\dT\d\d:\d\d:\d\d(Z|(\+|-)\d\d:\d\d)"/>.
    *
    * @param dt java.util.Date - the input date value to format.
    * @return String - the date value in Oagis DateTime format.
    */
   private final String getOagisDateTime(Date dt)
   {
      // Using RFC 822 4-digit time zone format
      DateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssZ");
      String sdf = null;
      StringBuffer dtBuf = new StringBuffer();

      //
      // Build the datetime in oagis format.  RFC 822 4-digit time zone format is used.
      // Eg: 2003-04-25T08:00:00-05:00.  See the javadocs for java.text.SimpleDateFormat.
      sdf = df.format(dt);

      dtBuf.append(sdf.substring(0, sdf.length() - 2));
      dtBuf.append(":");
      dtBuf.append(sdf.substring(sdf.length() - 2));

      df = null;
      sdf = null;

      return dtBuf.toString();
   }

   /**
    * Formats the input java.util.Date to return a String date value in the Oagis Date format
    *  e.g 2003-03-17.
    *
    * @param dt java.util.Date - the input date value to format.
    * @return String - the date value in Oagis Date format.
    */
   private final String getOagisDate(Date dt)
   {
      DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
      String sdf = df.format(dt);

      df = null;

      return sdf;
   }
   
   /**
    * Create the par file in the BOD/XML format
    * @return Par data in BOD format.
    * @throws SQLException 
    */
   private String parAsBod() throws SQLException
   {
      StringBuffer xml = new StringBuffer(1000);
      Date date = new Date();
      ResultSet rs = null;
      
      rs = m_ParData.executeQuery();
      
      try {
         //
         // Build prefix
         xml.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n");
         
         xml.append("<SyncPriceList \r\n");
         xml.append("xmlns=\"http://www.emeryonline.com/oagis\" \r\n");
         xml.append("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" \r\n");
         xml.append("xmlns:oa=\"http://www.openapplications.org/oagis\" \r\n");
         xml.append("xsi:schemaLocation=\"http://www.emeryonline.com/oagis \r\n");
         xml.append("http://www.emeryonline.com/oagis/Overlays/Emery/BODs/SyncPriceList.xsd\" \r\n");
         xml.append("revision=\"8.0\" ");
   
         xml.append("environment=\"");
         xml.append(OagConst.ENV[m_Env]);
         xml.append("\" lang=\"en-US\">\r\n");
   
         //
         // Build application area
         xml.append("<oa:ApplicationArea>\r\n");
         xml.append("<oa:CreationDateTime>");
         xml.append(getOagisDateTime(date));
         xml.append("</oa:CreationDateTime>\r\n");
   
         //
         // Build BOD id element
         xml.append("<oa:BODId>");
         xml.append(new GUID().toString());
         xml.append("</oa:BODId>\r\n");
   
         xml.append("</oa:ApplicationArea>\r\n");
         xml.append("<DataArea>\r\n");
   
         xml.append("<oa:Sync confirm=\"Never\">\r\n");
         xml.append("<oa:SyncCriteria>\r\n");
         xml.append("<oa:SyncExpression action=\"Change\"/>\r\n");
         xml.append("</oa:SyncCriteria>\r\n");
         xml.append("</oa:Sync>\r\n");
   
         xml.append("<oa:PriceList>\r\n");
         //
         // Build the header section.  This is done inside the lines loop, since need the value
         // of the par date to determine the price effective date, which lives in the header.      
         xml.append("<Header>\r\n");
         xml.append("<!-- Qualification of the intended audience of the price list e.g customer -->\r\n");
         xml.append("<oa:PriceListQualifier>\r\n");
         xml.append("<oa:Parties>\r\n");
         xml.append("<oa:CustomerParty>\r\n");
         xml.append("<oa:PartyId>\r\n");
         xml.append("<oa:Id>");
         xml.append(m_CustId);
         xml.append("</oa:Id>\r\n");
         xml.append("</oa:PartyId>\r\n");
         xml.append("</oa:CustomerParty>\r\n");
         xml.append("</oa:Parties>\r\n");
         xml.append("</oa:PriceListQualifier>\r\n");
         xml.append("<!-- Date price changes take effect -->\r\n");
         xml.append("<EffectiveDate>");
         xml.append(getOagisDate(date));
         xml.append("</EffectiveDate>\r\n");
         xml.append("</Header>\r\n");
      
         //
         // Add the lines
         while ( rs.next() ) {         
            xml.append("<Line>\r\n");            
            xml.append("<oa:ItemId>");
            xml.append("<oa:Id>");
            xml.append(rs.getString("item_id"));
            xml.append("</oa:Id>");
            xml.append("</oa:ItemId>\r\n");            
            xml.append("<oa:UnitPrice>\r\n");
            xml.append("<oa:Amount currency=\"USD\">");
            xml.append(rs.getDouble("sell"));
            xml.append("</oa:Amount>\r\n");
            xml.append("<oa:PerQuantity uom=\"");
            xml.append(rs.getString("ship_unit"));
            xml.append("\">1</oa:PerQuantity>\r\n");
            xml.append("</oa:UnitPrice>\r\n");            
            xml.append("<RetailPrice>\r\n");
            xml.append("<oa:Amount currency=\"USD\">");
            xml.append(rs.getDouble("retail"));
            xml.append("</oa:Amount>\r\n");
            xml.append("<oa:PerQuantity uom=\"");
            xml.append(rs.getString("ret_unit"));
            xml.append("\">1</oa:PerQuantity>\r\n");
            xml.append("</RetailPrice>\r\n");
            xml.append("</Line>\r\n");
         }
         
         xml.append("</oa:PriceList>\r\n");
         xml.append("</DataArea>\r\n");
         xml.append("</SyncPriceList>");
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      return xml.toString();
   }
   
   /**
    * 
    * @return PAR data in retail format.
    */
   private String parAsRet()
   {
      String par = "";
      return par;
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
            sql.append("select ");
            sql.append("   item.item_id, item.vendor_id, ");
            sql.append("   ship_unit.unit ship_unit, retail_unit.unit ret_unit, ");
            sql.append("   cust_procs.GETSELLPRICE(?, item.item_id) sell, ");
            sql.append("   cust_procs.GETRETAILPRICE(?, item.item_id) retail ");
            sql.append("from item ");
            sql.append("join ship_unit on ship_unit.unit_id = ship_unit_id ");
            sql.append("join retail_unit on retail_unit.unit_id = ret_unit_id ");
           
            if ( m_Dc > 0 )
               sql.append("join item_warehouse on item_warehouse.item_id = item.item_id and warehouse_id = ? ");
            
            sql.append("where ");
            sql.append("   item.item_type_id = 1 and in_catalog = 1 ");
            
            if ( m_VndId.length() > 0 )
               sql.append(String.format(" and item.vendor_id in (%s) ", m_VndId));
            
            sql.append("order by ");
            sql.append("   item.vendor_id, item.item_id");
            
            m_ParData = m_OraConn.prepareStatement(sql.toString());
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
         log.error("rocksolid.prepareStatements - null oracle connection");
      
      return isPrepared;
   }
   

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      int pcount = params.size();
      Param param = null;
      String tmp = null;
      
      try {
         for ( int i = 0; i < pcount; i++ ) {
            param = params.get(i);
            
            if ( param.name.equals("cust") )
               m_CustId = param.value;
            
            if ( param.name.equals("dc") )
               m_Dc = Integer.parseInt(param.value);
   
            if ( param.name.equals("env") )
               m_Env = Byte.parseByte(param.value);
                               
            if ( param.name.equals("fmt") )
               m_Format = Integer.parseInt(param.value);
            
            if ( param.name.equals("vnd") )
               m_VndId = param.value;
         }
         
         //
         // Set the file name extension based on the format.
         switch ( m_Format ) {
            case fmtBod: { 
               tmp = "SyncPriceList_%s_%s.xml";
               break;
            }
            
            default: {
               tmp = "par_%s_%s.txt";
               break;
            }
         }
         
         fileName.append(String.format(tmp, m_CustId, Long.toString(System.currentTimeMillis())));
         m_FileNames.add(fileName.toString());
      }
      
      finally {
         fileName = null;
         tmp = null;
         param = null;
      }
      
   }   
}
