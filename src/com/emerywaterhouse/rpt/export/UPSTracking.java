/**
 * File: UPSTracking.java
 * Description: Export file for UPS tracking numbers based on customer.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 05/18/2009
 * Last Update: $Id: UPSTracking.java,v 1.2 2014/06/12 16:29:01 jfisher Exp $
 *
 * History:
 *    $Log: UPSTracking.java,v $
 *    Revision 1.2  2014/06/12 16:29:01  jfisher
 *    Updated the po number field so "null" wouldn't show up.
 *
 *    Revision 1.1  2009/06/09 18:25:13  jfisher
 *    initial add
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class UPSTracking extends Report
{
   private final String ORDERID_ELEM  = "      <orderId>%d</orderId>\r\n";
   private final String CARTONID_ELEM = "      <tmsCartonId>%d</tmsCartonId>\r\n";
   private final String SHIPID_ELEM   = "      <shipId>%d</shipId>\r\n";
   private final String PONUM_ELEM    = "      <poNum>%s</poNum>\r\n";
   private final String TRACKNBR_ELEM = "      <trackingNbr>%s</trackingNbr>\r\n";
   private final String ACTFRT_ELEM   = "      <actualFrt>%f</actualFrt>\r\n";
   private final String DISCFRT_ELEM  = "      <discountedFrt>%f</discountedFrt>\r\n";
   private final String SHIPDATE_ELEM = "      <shipDate>%s</shipDate>\r\n";
   private final String WEIGHT_ELEM   = "      <weight>%f</weight>\r\n";
   private final String OVERSIZE_ELEM = "      <overSize>%d</overSize>\r\n";
   private final String SATDEL_ELEM =   "      <saturdayDelivery>%d</saturdayDelivery>\r\n";

   private PreparedStatement m_UpsData;
   private String m_CustId;
   private String m_Dc;
   private boolean m_Overwrite;

   /**
    *
    */
   public UPSTracking()
   {
      super();

      m_MaxRunTime = RptServer.HOUR * 12;
      m_CustId = "";
      m_Dc = "01";
      m_Overwrite = false;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_UpsData = null;
      m_Dc = null;
      m_CustId = null;

      super.finalize();
   }

   /**
    * Builds the output file based on the query selection criteria.  The output file is
    * XML based and contains the shipping information based on the customer.
    *
    * @return  boolean
    *    true if the file was created.
    *    false if there was some sort of error.
    */
   private boolean buildOutputFile()
   {
      SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
      StringBuffer line = new StringBuffer();
      FileOutputStream outFile = null;
      ResultSet upsData = null;
      long orderId = 0;
      long cartonId = 0;
      long shipId = 0;
      String trackingNbr = "";
      String shipDate = "";
      String poNum = "";
      double actFrt = 0.0;
      double discFrt = 0.0;
      double weight = 0.0;
      int overSize = 0;
      int satDeliv = 0;
      boolean result = false;

      try {
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         setCurAction("processing ups tracing information for customer: " + m_CustId);

         m_UpsData.setString(1, m_CustId);
         upsData = m_UpsData.executeQuery();

         line.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n");
         line.append("<upsTracking xmlns=\"http://www.emeryonline.com/ups\"\r\n");
         line.append("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\r\n");
         line.append("xsi:schemaLocation=\"http://www.emeryonline.com/ups ");
         line.append("http://www.emeryonline.com/ups/upstracking.xsd\">\r\n");

         while ( upsData.next() && m_Status == RptServer.RUNNING ) {
            orderId = upsData.getLong("order_id");
            cartonId = upsData.getLong("tms_carton_id");
            shipId = upsData.getLong("ship_id");
            trackingNbr = upsData.getString("tracking_nbr");
            actFrt = upsData.getDouble("actual_frt");
            discFrt = upsData.getDouble("discounted_frt");
            shipDate = fmt.format(upsData.getDate("ship_date"));
            weight = upsData.getDouble("weight");
            overSize = upsData.getInt("oversize");
            satDeliv = upsData.getInt("sat_delivery");
            poNum = upsData.getString("po_num");

            if ( poNum == null )
               poNum = "";

            line.append("   <shipment>\r\n");
            line.append(String.format(ORDERID_ELEM, orderId));
            line.append(String.format(CARTONID_ELEM, cartonId));
            line.append(String.format(SHIPID_ELEM, shipId));
            line.append(String.format(PONUM_ELEM, poNum));
            line.append(String.format(TRACKNBR_ELEM, trackingNbr));
            line.append(String.format(ACTFRT_ELEM, actFrt));
            line.append(String.format(DISCFRT_ELEM, discFrt));
            line.append(String.format(SHIPDATE_ELEM, shipDate));
            line.append(String.format(WEIGHT_ELEM, weight));
            line.append(String.format(OVERSIZE_ELEM, overSize));
            line.append(String.format(SATDEL_ELEM, satDeliv));
            line.append("   </shipment>\r\n");
         }

         line.append("</upsTracking>");
         outFile.write(line.toString().getBytes());

         result = true;
      }

      catch ( Exception ex ) {
         log.error("[UPS Tracking]", ex);
      }

      finally {
         line = null;
         fmt = null;
         shipDate = null;
         trackingNbr = null;
         poNum = null;

         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }

            catch ( IOException ex ) {
               ex.printStackTrace();
            }
         }
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_UpsData);
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
         log.fatal("[UPS Tracking]", ex);
      }

      finally {
         closeStatements();
         DbUtils.closeDbConn(m_EdbConn, null, null);

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Overrides the base class method to return the customer id the extract was run for.
    * @see com.emerywaterhouse.rpt.server.Report#getCustId()
    */
   @Override
   public String getCustId()
   {
      return m_CustId;
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
            sql.append("select order_header.order_id, order_header.po_num, tms_carton.* ");
            sql.append("from tms_carton ");
            sql.append("join shipment_order on shipment_order.ship_id = tms_carton.ship_id ");
            sql.append("join order_header on order_header.order_id = shipment_order.order_id and ");
            sql.append("   order_header.customer_id = ? ");
            sql.append("where ship_date = trunc(now()) ");
            sql.append("order by order_header.order_id, tms_carton_id");

            m_UpsData = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[UPS Tracking]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("ups tracking export: missing database connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      String fmt1 = "%d-%s-upstracking-%s.xml";
      String fmt2 = "%s-upstracking-%s.xml";
      StringBuffer fileName = new StringBuffer();
      int pcount = params.size();
      Param param = null;

      try {
         for ( int i = 0; i < pcount; i++ ) {
            param = params.get(i);

            if ( param.name.equals("dc") )
               m_Dc = param.value;

            if ( param.name.equals("cust") )
               m_CustId = param.value;

            if ( param.name.equals("overwrite") )
               m_Overwrite = param.value.equalsIgnoreCase("true") ? true : false;
         }

         //
         // Some customers want the same file name each time.  If that's the case, we
         // need to overwrite what we have.
         if ( m_Overwrite )
            fileName.append(String.format(fmt2, m_CustId, m_Dc));
         else
            fileName.append(String.format(fmt1, System.currentTimeMillis(), m_CustId, m_Dc));

         m_FileNames.add(fileName.toString());
      }

      finally {
         fileName = null;
         param = null;
         fmt1 = null;
         fmt2 = null;
      }
   }
}
