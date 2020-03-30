/**
 * File: ActiveWhseItemReport.java
 * Description: Generates a spreadsheet report of all active,
 * non-virtual, items.  Requested by Tom Poole.
 *
 * @author Erik Pearson
 *
 * Create Date: 08/24/2010
 * Last Update: $Id: ActiveWhseItemReport.java,v 1.2 2013/11/15 15:14:02 jfisher Exp $
 *
 * History
 *    $Log: ActiveWhseItemReport.java,v $
 *    Revision 1.2  2013/11/15 15:14:02  jfisher
 *    Removed the in catalog check at the item level.
 *
 *    Revision 1.1  2010/08/29 00:28:49  epearson
 *    Initial add
 *
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.sql.PreparedStatement;
import java.sql.SQLException;

import com.emerywaterhouse.rpt.server.SimpleReport;

public class ActiveWhseItemReport extends SimpleReport {

	/**
	 * Constructor
	 */
	public ActiveWhseItemReport() {
		super();
		this.setReportTitle("Active Warehouse Items Report");
	}

	/**
	 * Builds the filename based on the system time and UID.
	 * @see com.emerywaterhouse.rpt.server.SimpleReport
	 *
	 * @return	the report filename
	 */
	@Override
	protected String buildFileName() {
		String currentTime = Long.toString(System.currentTimeMillis()).substring(3);

		return currentTime + "-" + m_RptProc.getUid() + "-item_rpt.xls";
	}

	/**
	 * Builds the query for the item data.
	 * @see com.emerywaterhouse.rpt.server.SimpleReport
	 *
	 * @return	the item data query
	 * @throws	a sql exception that is caught in SimpleReport.createReport()
	 */
	@Override
	protected PreparedStatement buildReportQuery() throws SQLException {
		StringBuffer sql = new StringBuffer(1024);

		sql.append("select item.item_id, item_upc.upc_code, ");
		sql.append("item.description, vendor.name as VENDOR_NAME, item_price.sell, ");
		sql.append("retail_c, ship_unit.unit, item.stock_pack, ");
		sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', null, 'N') NBC, ");
		sql.append("mis.unit_sales, mis.unit_dollars, item.flc_id flc, ");
		sql.append("flc_procs.GETNRHAID(item.flc_id) nrha, ");
		sql.append("to_char(setup_date, 'yyyy/MM/dd') setup_date, ");
		sql.append("itemwhse.dc_names ");

		sql.append("from item ");
		sql.append("join vendor on (vendor.vendor_id = item.vendor_id) ");
		sql.append("left outer join item_price on (item_price.item_id = item.item_id) ");
		sql.append("join ship_unit on (ship_unit.unit_id = item.ship_unit_id) ");
		sql.append("left outer join item_upc on (item_upc.item_id = item.item_id) ");
		sql.append("join item_disp on (item_disp.disp_id = item.disp_id) ");
		sql.append("join item_type on (item_type.item_type_id = item.item_type_id) ");
		sql.append("join broken_case on (broken_case.broken_case_id = item.broken_case_id) ");

		// Get 12 months of unit sales, have to go back 13 months because
		// there is no data for the current month
		sql.append("left outer join ");
		sql.append("(select item_nbr, sum(units_shipped) unit_sales, sum(dollars_shipped) unit_dollars ");
		sql.append("from sa.monthlyitemsales where ");
		sql.append("months_between(last_day(sysdate), to_date(year_month, 'yyyymm')) <= 13 ");
		sql.append("group by item_nbr) ");
		sql.append("mis on (mis.item_nbr = item.item_id) ");

		// Get warehouse location
		sql.append("join ");
		sql.append("(select item_warehouse.item_id, wmsys.wm_concat(warehouse.name) dc_names ");
		sql.append("from item_warehouse, warehouse ");
		sql.append("where item_warehouse.warehouse_id = warehouse.warehouse_id ");
		sql.append("group by item_warehouse.item_id) ");
		sql.append("itemwhse on (itemwhse.item_id = item.item_id) ");

		sql.append("where "); // Only active non virtual items
		sql.append("item_type.itemtype = 'STOCK' and ");
		sql.append("item_disp.disposition in ('BUY-SELL','NOBUY') and ");
		sql.append("item_upc.primary_upc = 1 and ");
		sql.append("item_price.sell_date = (select max(sell_date) from item_price ");
		sql.append("where sell_date <= trunc(sysdate) and item_id = item.item_id) ");

		sql.append("order by item.item_id asc ");

		return m_OraConn.prepareStatement(sql.toString());
	}
}
