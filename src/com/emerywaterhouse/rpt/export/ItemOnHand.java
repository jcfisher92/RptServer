/**
 * Title:			ItemOnHand.java
 * Description:
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date:	Dec 29, 2009
 * Last Update:   $Id: ItemOnHand.java,v 1.4 2012/07/11 17:30:56 jfisher Exp $
 * <p>
 * History:
 *		$Log: ItemOnHand.java,v $
 *		Revision 1.4  2012/07/11 17:30:56  jfisher
 *		in_catalog modification
 *
 *		Revision 1.3  2011/01/11 19:30:47  prichter
 *		Added a check for report status so it can be stopped
 *
 *		Revision 1.2  2010/01/25 01:05:33  prichter
 *		Retail web project.
 *
 *		Revision 1.1  2010/01/03 10:27:05  prichter
 *		Production versions
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import com.emerywaterhouse.oag.build.bod.UpdateInventoryCount;
import com.emerywaterhouse.oag.build.noun.InventoryCount;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class ItemOnHand extends Report {

   private static final String WHERE_ITEM = "and item_entity_attr.item_id = '%s'\r\n";
   private static final String FLC_JOIN   = "   and flc.flc_id = '%s' \r\n";
   private static final String MDC_JOIN   = "   and mdc.mdc_id = '%s' \r\n";
   private static final String NRHA_JOIN  = "   and nrha.nrha_id = '%s'\r\n";

   private String m_Warehouse;	// The warehouse the report is run for
   private String m_FacilityId;	// The fascor facility id
   private String m_DataFmt;     // The output format, xml, flat, excel
   private String m_DataSrc;     // The data source to limit the data by: flc, mdc, nrha, item
   private String m_SrcId;       // The value of the data source limiter.
   private boolean m_Overwrite;	// If false, a unique file name will be created

   private PreparedStatement m_Items;

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
         if ( m_DataFmt.equals("xml") )
            result = buildXml(outFile);
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
    * Builds the on hand quantity export in XML format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildXml(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      ResultSet rs = null;
      UpdateInventoryCount doc = new UpdateInventoryCount();
      InventoryCount inv = null;
      InventoryCount.Line item = null;

      m_Items.setString(1, m_Warehouse);
      rs = m_Items.executeQuery();

      inv = doc.addInventoryCount();

      try {
      	while ( rs.next() && m_Status == RptServer.RUNNING ) {
      		setCurAction("Processing item " + rs.getString("item_id"));
      		item = inv.addLine();
      		item.setItemId(inv.getPrefix(), rs.getString("item_id"));
      		item.setQuantity(inv.getPrefix(), rs.getInt("qty"));
      	}

      	if ( m_Status == RptServer.RUNNING ) {
      		outFile.write(doc.toString().getBytes());
      		result = true;
      	}

      	else
      		result = false;
      }

      finally {
         setCurAction(String.format("finished processing on hand data"));
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return result;
   }

   /**
    * Close prepared statements and connections
    */
   private void closeStatements()
   {
   	DbUtils.closeDbConn(m_EdbConn, m_Items, null);
   	m_Items = null;
      m_EdbConn = null;
   }

	/* (non-Javadoc)
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
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer();
      Statement stmt = null;
      ResultSet rs = null;


      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            // Find the fascor facility id.  This is needed by the item
         	// query to obtain the current on hand quantities
         	stmt = m_EdbConn.createStatement();
            rs = stmt.executeQuery("select * from warehouse where name = '" + m_Warehouse + "'");

            if ( rs.next() )
            	m_FacilityId = rs.getString("fas_facility_id");

            DbUtils.closeDbConn(null, stmt, rs);
            rs = null;
            stmt = null;

            sql.append("select ");
            sql.append("    item_entity_attr.item_id, ejd_item_warehouse.qoh qty ");
            sql.append("from ");
            sql.append("   item_entity_attr ");
            sql.append("join ejd_item on item_entity_attr.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("join warehouse on warehouse.name = ? ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ");
            sql.append("                       ejd_item_warehouse.warehouse_id = warehouse.warehouse_id and ejd_item_warehouse.in_catalog = 1 ");
            sql.append("join bmi_item on bmi_item.item_id = item_entity_attr.item_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append(m_DataSrc.equals("flc") ? String.format(FLC_JOIN, m_SrcId) : "\r\n");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append(m_DataSrc.equals("mdc") ? String.format(MDC_JOIN, m_SrcId) : "\r\n");
            sql.append("join nrha on nrha.nrha_id = mdc.nrha_id ");
            sql.append(m_DataSrc.equals("nrha") ? String.format(NRHA_JOIN, m_SrcId) : "\r\n");
            sql.append("where ");
            sql.append("   item_entity_attr.item_type_id = 1 ");
            sql.append( m_DataSrc.equals("item") ? String.format(WHERE_ITEM, m_SrcId) : "\r\n");
            sql.append("order by item_entity_attr.item_id");
            m_Items = m_EdbConn.prepareStatement(sql.toString());

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
         log.error("CustomerPriceList.prepareStatements - null EDB connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      String tmp = Long.toString(System.currentTimeMillis());
      int pcount = params.size();
      Param param = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("datafmt") )
            m_DataFmt = param.value;

         if ( param.name.equals("datasrc") )
            m_DataSrc = param.value;

         if ( param.name.equals("srcid") )
            m_SrcId = param.value;

         if ( param.name.equals("overwrite") )
            m_Overwrite = param.value.equalsIgnoreCase("true") ? true : false;

         if ( param.name.equals("warehouse" ) )
         	m_Warehouse = param.value;
      }

      if ( m_DataSrc == null )
      	m_DataSrc = "";

      if ( m_SrcId == null )
      	m_SrcId = "";

      //
      // Some customers want the same file name each time.  If that's the case, we
      // need to overwrite what we have.
      if ( !m_Overwrite ) {
         fileName.append(tmp);
         fileName.append("-");

         if ( m_DataSrc != null && m_DataSrc.trim().length() > 0 ) {
            fileName.append(m_DataSrc);
            fileName.append("-");

            if ( m_SrcId != null && m_SrcId.trim().length() > 0 ) {
	            fileName.append(m_SrcId);
	            fileName.append("-");
            }
         }
      }

      fileName.append("emery-itemqoh.xml");
      m_FileNames.add(fileName.toString());
   }

}

