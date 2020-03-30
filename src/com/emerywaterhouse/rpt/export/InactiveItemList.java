/**
 * Title:			NonCatalogItems.java
 * Description:
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date:	Jan 11, 2010
 * Last Update:   $Id: InactiveItemList.java,v 1.4 2012/07/11 17:30:56 jfisher Exp $
 * <p>
 * History:
 *		$Log: InactiveItemList.java,v $
 *		Revision 1.4  2012/07/11 17:30:56  jfisher
 *		in_catalog modification
 *
 *		Revision 1.3  2011/01/11 19:30:47  prichter
 *		Added a check for report status so it can be stopped
 *
 *		Revision 1.2  2010/01/30 01:38:02  prichter
 *		Exclude NOBUY items if they're still in the catalog
 *
 *		Revision 1.1  2010/01/25 01:05:33  prichter
 *		Retail web project.
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class InactiveItemList extends Report {

	private String m_DataFmt;

	/**
	 * Creates the output file
	 *
	 * @return boolean - true if successful
	 * @throws FileNotFoundException
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

	private boolean buildXml(FileOutputStream outFile) throws Exception
	{
		boolean result = false;
		StringBuffer xml = new StringBuffer();
		StringBuffer sql = new StringBuffer();
		Statement stmt = null;
		ResultSet rs = null;

		sql.append("select item.item_id ");
		sql.append("from item ");
		sql.append("join item_disp on item_disp.disp_id = item.disp_id ");
		sql.append("join item_type on item_type.item_type_id = item.item_type_id ");
		sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id and ");
		sql.append("                  ship_unit.unit <> 'AST' ");
		sql.append("join warehouse on warehouse.name = 'PORTLAND' ");
		sql.append("left join item_warehouse on item_warehouse.warehouse_id = warehouse.warehouse_id and ");
		sql.append("                            item_warehouse.item_id = item.item_id ");
		sql.append("where item_warehouse.in_catalog = 0 or ");
		sql.append("      item_disp.disposition not in ('BUY-SELL','NOBUY') or ");
		sql.append("      item_type.itemtype <> 'STOCK' or ");
		sql.append("      not exists(select item_warehouse.warehouse_id ");
		sql.append("                 from item_warehouse ");
		sql.append("                 where item_warehouse.item_id = item.item_id and ");
		sql.append("                       item_warehouse.warehouse_id = warehouse.warehouse_id) or ");
		sql.append("      item_warehouse.active = 0 or ");
		sql.append("      not exists(select item_id from bmi_item where bmi_item.item_id = item.item_id)  ");

		stmt = m_OraConn.createStatement();
		rs = stmt.executeQuery(sql.toString());

		try {
			xml.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
			xml.append("<InactiveItems>\r\n");

			while ( rs.next() && m_Status == RptServer.RUNNING ) {
				xml.append("<ItemId>" + rs.getString("item_id") + "</ItemId>\r\n");
			}

			xml.append("</InactiveItems>\r\n");

			if ( m_Status == RptServer.RUNNING ) {
				outFile.write(xml.toString().getBytes());
				result = true;
			}

			else
				result = false;
		}

		finally {
			DbUtils.closeDbConn(m_OraConn, stmt, rs);
			rs = null;
			stmt = null;
			m_OraConn = null;
		}

		return result;
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
         m_OraConn = m_RptProc.getOraConn();
         if ( m_OraConn != null )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }

      finally {
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
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
      }

      fileName.append(tmp);
      fileName.append("-");
      fileName.append("non-catalog-items.xml");
      m_FileNames.add(fileName.toString());
   }

}

