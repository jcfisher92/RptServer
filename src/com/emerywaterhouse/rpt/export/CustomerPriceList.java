/**
 * Title:			CustomerPriceList.java
 * Description:
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date:	Dec 29, 2009
 * Last Update:   $Id: CustomerPriceList.java,v 1.10 2013/03/12 17:58:45 prichter Exp $
 * <p>
 * History:
 *		$Log: CustomerPriceList.java,v $
 *		Revision 1.10  2013/03/12 17:58:45  prichter
 *		Added a reqsource parameter to allow special processing depending on why the extract was requested.  A source of retailweb will prevent displays from being included.
 *
 *		Revision 1.9  2012/07/11 17:30:56  jfisher
 *		in_catalog modification
 *
 *		Revision 1.8  2011/01/11 19:30:47  prichter
 *		Added a check for report status so it can be stopped
 *
 *		Revision 1.7  2010/03/23 15:07:22  prichter
 *		Add the customer id to the file name
 *
 *		Revision 1.6  2010/02/23 06:55:28  prichter
 *		Added handling for retail customization passed as parameters.  This option currently only handles retails A, B, C, and D.  The hooks were added for more options that may need to be added eventually.
 *
 *		Revision 1.5  2010/02/16 02:32:19  prichter
 *		Change the join to price_sensitivity to a left outer join so manually priced and image items aren't skipped.
 *
 *		Revision 1.4  2010/02/14 09:49:18  prichter
 *		Added customized retail capabilities
 *
 *		Revision 1.3  2010/02/01 22:52:31  prichter
 *		Modified main query to handle customers with only contract pricing
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
import java.util.ArrayList;

import com.emerywaterhouse.oag.build.bod.SyncPriceList;
import com.emerywaterhouse.oag.build.noun.PriceList;
import com.emerywaterhouse.oag.build.noun.PriceList.ItemPrice;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class CustomerPriceList extends Report {
   private static final String WHERE_ITEM = "and item_entity_attr.item_id = '%s' and \r\n";
   private static final String FLC_JOIN   = "   and flc.flc_id = '%s' \r\n";
   private static final String MDC_JOIN   = "   and mdc.mdc_id = '%s' \r\n";
   private static final String NRHA_JOIN  = "   and nrha.nrha_id = '%s'\r\n";

	private String m_CustId;		// The customer the prices calculated for
   private String m_DataFmt;     // The output format, xml, flat, excel
   private String m_DataSrc;     // The data source to limit the data by: flc, mdc, nrha, item
   private String m_SrcId;       // The value of the data source limiter.
   private boolean m_Overwrite;	// If false, a unique file name will be created
   private String m_ReqSource;	// The source of this request to allow for special processing

   private PreparedStatement m_Items;

   private ArrayList<RetailOption> m_ItemRetail;
   private ArrayList<RetailOption> m_ImageRetail;
   private ArrayList<RetailOption> m_FlcRetail;
   private ArrayList<RetailOption> m_NrhaRetail;
   private ArrayList<RetailOption> m_StoreRetail;

   private void addRetailOption(String param) throws Exception
   {
   	RetailOption opt = new RetailOption(param);

   	if ( opt.getOption().equalsIgnoreCase("item")) {
   		if ( m_ItemRetail == null )
   			m_ItemRetail = new ArrayList<RetailOption>();

			m_ItemRetail.add(opt);
   	}

   	if ( opt.getOption().equalsIgnoreCase("image") ) {
   		if ( m_ImageRetail == null )
   			m_ImageRetail = new ArrayList<RetailOption>();

   		m_ImageRetail.add(opt);
   	}

   	if ( opt.getOption().equalsIgnoreCase("flc")) {
   		if ( m_FlcRetail == null )
   			m_FlcRetail = new ArrayList<RetailOption>();

			m_FlcRetail.add(opt);
   	}

   	if ( opt.getOption().equalsIgnoreCase("nrha") ) {
   		if ( m_NrhaRetail == null )
   			m_NrhaRetail = new ArrayList<RetailOption>();

   		m_NrhaRetail.add(opt);
   	}

   	if ( opt.getOption().equalsIgnoreCase("store")) {
   		if ( m_StoreRetail == null )
   			m_StoreRetail = new ArrayList<RetailOption>();

   		m_StoreRetail.add(opt);
   	}
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
    * Builds the customer retail price export in XML format.
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
      SyncPriceList doc = new SyncPriceList();
      PriceList price = null;
      ItemPrice item = null;
      double retail;

      m_Items.setString(1, m_CustId);
      rs = m_Items.executeQuery();

		price = doc.addDocument();
		price.setBusinessId(m_CustId);
		price.setBusinessType("customer");

      try {
      	while ( rs.next() && m_Status == RptServer.RUNNING ) {
      		setCurAction("Processing customer " + m_CustId + " item " + rs.getString("item_id"));
      		retail = calcRetail(rs);

      		if ( retail > 0.01 ) {
	      		item = price.addItem();
	      		item.setItemId(rs.getString("item_id"));
	      		item.setSellPrice(rs.getDouble("cost"));
	      		item.setSellPerQuantity(1, rs.getString("ship_unit"));
	      		item.setRetailPrice(retail);
	      		item.setRetailPerQuantity(1, rs.getString("retail_unit"));
	      		item.setSellPack(rs.getInt("stock_pack"));
	      		item.setSellPackUom(rs.getString("ship_unit"));
	      		item.setRetailPack(rs.getInt("retail_pack"));
	      		item.setRetailPackUom(rs.getString("retail_unit"));
      		}
      	}

      	if ( m_Status == RptServer.RUNNING ) {
      		outFile.write(doc.toString().getBytes());
      		result = true;
      	}

      	else
      		result = false;
      }

      finally {
         setCurAction(String.format("finished processing retail price data"));
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return result;
   }

   /**
    * Calculates the customer's desired retail based on report options.  If there
    * are no options that cover the current item, the customer's regularly
    * calculated retail will be used.  This was added for the Retail Web project.
    * It allows customers to have different retails on the retail web site from
    * there regular store retails.
    *
    * @param opt RetailOption - the retail options loaded from the report parameters
    * @param rs ResultSet - the ResultSet containing the customer's price data
    * @return double - the retail price that will be included on the report
    * @throws Exception
    */
   private double calcCustomRetail(RetailOption opt, ResultSet rs) throws Exception
   {
   	double retail = 0;

		retail = rs.getDouble("retail");

		if ( opt.getPriceOption().equalsIgnoreCase("a") )
			retail = rs.getDouble("retail_a");

		if ( opt.getPriceOption().equalsIgnoreCase("b") )
			retail = rs.getDouble("retail_b");

		if ( opt.getPriceOption().equalsIgnoreCase("c") )
			retail = rs.getDouble("retail_c");

		if ( opt.getPriceOption().equalsIgnoreCase("d") )
			retail = rs.getDouble("retail_d");

		//TODO other options
		//If and when other options are needed, add them here

		return retail;
   }

   /**
    * Returns the retail price that will be included on the report.  This is included
    * for the Retail Web site.  Customers may want different retails on the web
    * site than they normally use in their store.  The default value is the
    * retail calculated by the main query.
    *
    * @param rs ResultSet - the ResultSet from the main pricing query.
    * @return double - the customer's retail price.
    */
   private double calcRetail(ResultSet rs)
   {
   	double retail = 0;
   	String item = null;
   	RetailOption opt = null;

   	try {
   		item = rs.getString("item_id");
   		retail = rs.getDouble("retail");
   		opt = getRetailOption(rs);

   		if ( opt != null ) {
   			retail = calcCustomRetail(opt, rs);
   		}
   	}

   	catch ( Exception e ) {
  			log.error("Unable to calculate retail for " + m_CustId + " " + item);
   	}

   	return retail;
   }

   /**
    * Resource cleanup
    */
   private void closeStatements()
   {
   	DbUtils.closeDbConn(null, m_Items, null);
   	m_Items = null;

   	m_ItemRetail = null;
   	m_ImageRetail = null;
   	m_FlcRetail = null;
   	m_StoreRetail = null;
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
	 * Returns a retail options based on FLC.  If none exists, returns null.
	 *
	 * @param rs ResultSet - the result set from PreparedStatement m_Items
	 * @return RetailOption - the retail option based on the current item's flc
	 */
	private RetailOption getFlcRetailOption(ResultSet rs)
	{
		RetailOption opt = null;
		//TODO getFlcRetailOption
		return opt;
	}

	/**
	 * Returns a retail options based on NRHA.  If none exists, returns null.
	 *
	 * @param rs ResultSet - the result set from PreparedStatement m_Items
	 * @return RetailOption - the retail option based on the current item's nrha
	 */
	private RetailOption getNrhaRetailOption(ResultSet rs)
	{
		RetailOption opt = null;
		//TODO getNrhaRetailOption
		return opt;
	}
	/**
	 * Returns a retail options for image items  If none exists, returns null.
	 *
	 * @param rs ResultSet - the result set from PreparedStatement m_Items
	 * @return RetailOption - the retail option for image items
	 */
	private RetailOption getImageRetailOption(ResultSet rs)
	{
		RetailOption opt = null;
		//TODO getImageRetailOption
		return opt;
	}

	/**
	 * Returns a retail options for a specific item.  If none exists, returns null.
	 *
	 * @param rs ResultSet - the result set from PreparedStatement m_Items
	 * @return RetailOption - the retail option for a specific item
	 */
	private RetailOption getItemRetailOption(ResultSet rs)
	{
		RetailOption opt = null;
		//TODO getItemRetailOption
		return opt;
	}

	/**
	 * Returns the default retail for the store.  If none exists, returns null.
	 *
	 * @param rs ResultSet - the result set from PreparedStatement m_Items
	 * @return RetailOption - the default retail option for the store
	 */
	private RetailOption getStoreRetailOption(ResultSet rs) throws Exception
	{
		RetailOption opt = null;

		if ( m_StoreRetail == null )
			return opt;

		for ( int i = 0; i < m_StoreRetail.size(); i++ ) {
			if ( m_StoreRetail.get(i).getSensitivity().length() > 0 &&
				  m_StoreRetail.get(i).getSensitivity().equals(rs.getString("sen_code_id")) )
			  opt = m_StoreRetail.get(i);
		}

		if ( opt == null ) {
			for ( int i = 0; i < m_StoreRetail.size(); i++ ) {
				if ( m_StoreRetail.get(i).getSensitivity().length() == 0 ||
					  m_StoreRetail.get(i).getSensitivity().equalsIgnoreCase("all") )
					opt = m_StoreRetail.get(i);
			}
		}

		return opt;
	}

	/**
	 * Searches for a customized retail option for the current item.  The options
	 * are passed to the report as parameters.  No all options have been implemented
	 * so check before using one.
	 *
	 * @param rs RetailSet - The ResultSet from PreparedStatement m_Items
	 * @return RetailOption - the retail options that applies to the current item
	 */
	private RetailOption getRetailOption(ResultSet rs)
	{
		RetailOption opt = null;
		String item = null;

		try {
			item = rs.getString("item_id");

			opt = getItemRetailOption(rs);

			if ( opt == null )
				opt = getImageRetailOption(rs);

			if ( opt == null )
				opt = getFlcRetailOption(rs);

			if ( opt == null )
				opt = getNrhaRetailOption(rs);

			if ( opt == null )
				opt = getStoreRetailOption(rs);
		}
		catch ( Exception e ) {
			log.error("Unable to get retail option for " + m_CustId + " " + item);
			log.error("exception", e);
		}

		return opt;
	}

   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;
      String shipUnitCondition = " ";
      
      if ( m_ReqSource != null )
         shipUnitCondition = m_ReqSource.equalsIgnoreCase("retailweb") ? " and ship_unit.unit <> 'DSP' " : " "; 
      	

      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("   item_entity_attr.item_id, ejd_item_warehouse.stock_pack, item_entity_attr.retail_pack, ");
            sql.append("   ship_unit.unit ship_unit, retail_unit.unit retail_unit,  ");
            sql.append("   (select price from  ejd_cust_procs.get_sell_price(customer.customer_id, item_entity_attr.item_ea_id)) as cost, ");
            sql.append("   ejd_price_procs.get_retail_price(customer.customer_id, item_entity_attr.item_ea_id) as retail, ");
            sql.append("   ejd_item_price.buy, ejd_item_price.sell,  ");
            sql.append("   ejd_item_price.retail_a, ejd_item_price.retail_b, ejd_item_price.retail_c, ");
            sql.append("   nvl(ejd_item_price.retail_d, ejd_item_price.retail_c) retail_d, ");
            sql.append("   item_price_method.price_method, price_sensitivity.sen_code_id ");
            sql.append("from        ");
            sql.append("   item_entity_attr ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.in_catalog = 1  ");
            sql.append("join customer on customer.customer_id = ? ");
            sql.append("join cust_warehouse on cust_warehouse.customer_id = customer.customer_id and ");
            sql.append("                       cust_warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");
            sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id " + shipUnitCondition);
            sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("      and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id  ");
            sql.append("join item_price_method on item_price_method.method_id = ejd_item_price.method_id ");
            sql.append("left outer join price_sensitivity on price_sensitivity.sen_code_id = ejd_item_price.sen_code_id ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");  
            sql.append("join flc on flc.flc_id = ejd_item.flc_id  ");
            sql.append(m_DataSrc.equals("flc") ? String.format(FLC_JOIN, m_SrcId) : "\r\n");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id  ");
            sql.append(m_DataSrc.equals("mdc") ? String.format(MDC_JOIN, m_SrcId) : "\r\n");
            sql.append("join nrha on nrha.nrha_id = mdc.nrha_id ");
            sql.append(m_DataSrc.equals("nrha") ? String.format(NRHA_JOIN, m_SrcId) : "\r\n");
            sql.append("where ");
            sql.append("   item_entity_attr.item_type_id = 1 and ");
            sql.append( m_DataSrc.equals("item") ? String.format(WHERE_ITEM, m_SrcId) : "\r\n");
            sql.append("   (exists (select * from cust_price_method ");
            sql.append("            join price_method on price_method.price_method_id = cust_price_method.price_method_id and ");
            sql.append("                                 price_method.description = 'RETAIL RIGHT' ");
            sql.append("            where cust_price_method.customer_id = customer.customer_id) or   ");
            sql.append("   ejd.cust_procs.pricing_check(customer.customer_id, item_entity_attr.item_id) = 'OK') ");
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
         log.error("CustomerPriceList.prepareStatements - null oracle connection");

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

         if ( param.name.equals("cust") )
            m_CustId = param.value.trim();

         if ( param.name.equals("datafmt") )
            m_DataFmt = param.value.trim();

         if ( param.name.equals("datasrc") )
            m_DataSrc = param.value.trim();

         if ( param.name.equals("srcid") )
            m_SrcId = param.value.trim();

         if ( param.name.equals("overwrite") )
            m_Overwrite = param.value.trim().equalsIgnoreCase("true") ? true : false;

         if ( param.name.equals("retail.option") ) {
         	try {
         		addRetailOption(param.value.trim());
         	}
         	catch ( Exception e ) {
         		log.error("Error parsing retail option " + param.value);
         		log.error("exception", e);
         	}
         }
         
         if ( param.name.equals("reqsource") ) {
         	m_ReqSource = param.value.trim();
         }
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
         fileName.append(m_CustId);
         fileName.append("-");

         if ( m_DataSrc != null && m_DataSrc.trim().length() > 0 ) {
            fileName.append(m_DataSrc.trim());
            fileName.append("-");

            if ( m_SrcId != null && m_SrcId.trim().length() > 0 ) {
	            fileName.append(m_SrcId.trim());
	            fileName.append("-");
            }
         }
      }

      fileName.append("emery-pricelist.xml");
      m_FileNames.add(fileName.toString());
   }

   /**
    * An inner class to represent a pricing option passed as a parameter.  This
    * is used for customers who want to override their normal store pricing
    * structure in this report.  It was created for the RetailWeb site so that
    * customer can have different retail on the web site from those in the store.
    *
    * @author prichter
    *
    */
   public class RetailOption
   {
   	private String m_Option;
   	private String m_OptionValue;
   	private String m_Sensitivity;
   	private String m_PriceOption;

   	/**
   	 * Default Constructor
   	 */
   	public RetailOption()
   	{
   		m_Option = "";
   		m_OptionValue = "";
   		m_Sensitivity = "";
   		m_PriceOption = "";
   	}

   	/**
   	 * Constructor with the retail parameter value string.  Parses the parameter
   	 * string.
   	 *
   	 * @param param String - the parameter values passed to the report
   	 * @throws Exception
   	 */
   	public RetailOption(String param) throws Exception
   	{
   		this();
   		parse(param);
   	}

   	/**
   	 * Returns the retail option.  Options are Store, Image, Nrha, Flc, and Item.
   	 *
   	 * @return String - the retail option.  Defines the scope of the OptionValue
   	 * 	parameter.  E.g., if the option = 'FLC', the option value will be a
   	 * 	flc_id.
   	 */
   	public String getOption()
   	{
   		return m_Option;
   	}

   	/**
   	 * The value of the option.  E.g., if the option is 'Nrha', the value will
   	 * be an nrha_id.
   	 *
   	 * @return String - the option value
   	 */
   	public String getOptionValue()
   	{
   		return m_OptionValue;
   	}

   	/**
   	 * Defines how the retail price will be calculated.  Valid values include:
   	 * 'a', 'b', 'c', or 'd' for the 4 Emery suggested retails, 'margin',
   	 * 'variance', and 'price' (fixed price).
   	 *
   	 * NOTE: only 'a','b','c', and 'd' are currently implemented.  The others
   	 * will need to be added if they are to be used.
   	 *
   	 * @return String - the price options
   	 */
   	public String getPriceOption()
   	{
   		return m_PriceOption;
   	}

   	/**
   	 * Sets the sensitivity code this price option applies to.
   	 *
   	 * @return String - the sensitivity code this price applies to.
   	 */
   	public String getSensitivity()
   	{
   		return m_Sensitivity;
   	}

   	/**
   	 * Parses the option string.  The options are passed to the report as
   	 * parameters with the form:
   	 *
   	 *   option,optvalue,sensitivity,priceoption
   	 *
   	 * For example:
   	 *
   	 * 	store,,,c
   	 *
   	 * - the option applies to the entire store
   	 * - the option value is ignored	because it's the entire store
   	 * - it applies to all sensitivity codes
   	 * - use Emery Retail C
   	 *
   	 * 	store,,8,d
   	 *
   	 * - entire store
   	 * - no value option
   	 * - sensitivity 8
   	 * - Retail D
   	 *
   	 * 	nrha,01,,a
   	 *
   	 * - Option nrha
   	 * - nrha 01
   	 * - no sensitivity
   	 * - Retail A
   	 *
   	 * @param val - the parameter value
   	 * @throws Exception - if the string can't be parsed
   	 */
   	public void parse(String val) throws Exception
   	{
   		int i, j;

   		i = val.indexOf(',');

   		if ( i == -1 )
   			m_Option = "store";
   		else
   			m_Option = val.substring(0, i);

   		if ( i != -1 ) {
   			j = i + 1;

   			i = val.indexOf(',', j);

   			if ( i != -1 )
   				m_OptionValue = val.substring(j, i);
   		}

   		if ( i != -1 ) {
   			j = i + 1;

   			i = val.indexOf(',', j);

   			if ( i != -1 )
   				m_Sensitivity = val.substring(j, i);
   		}

   		if ( i < val.length() )
   			m_PriceOption = val.substring(i + 1);
   	}

   	/**
   	 * Sets the retail item selection option.
   	 *
   	 * NOTE:  Not all options have been implemented.  Check the code to make
   	 * sure the option you're trying to add exists.
   	 *
   	 * @param val String - the retail item selection option.
   	 * 	Valid values are: 'Store', 'Nrha', 'Flc', 'Image', 'Item'.
   	 */
   	public void setOption(String val)
   	{
   		m_Option = val;
   	}

   	/**
   	 * Sets the value associated with the item selection option.
   	 * For example, the flc_id, nrha_id, or item_id.
   	 *
   	 * @param val String - the option value
   	 */
   	public void setOptionValue(String val)
   	{
   		m_OptionValue = val;
   	}

   	/**
   	 * The option used to calculate the retails.  These follow the same basic
   	 * rules as EIS CRP (see cust_procs.crp_update and cust_procs.getRetailPrice),
   	 * except that the initial price option cannot be used tp establish a margin
   	 * setting.
   	 *
   	 * NOTE: The only options that are currently implemented are Emery
   	 * 	retail A, B, C, and D.  Others will be added as needed.
   	 *
   	 * @param val String - the Retail option
   	 */
   	public void setPriceOption(String val)
   	{
   		m_PriceOption = val;
   	}

   	/**
   	 * Sets the sensitivity code.  If blank or null, this option will apply
   	 * to all sensitivities.  Otherwise, it will only apply to items priced using
   	 * a matching sensitivity.
   	 *
   	 * @param val String - the sensitivity code
   	 */
   	public void setSensitivity(String val)
   	{
   		m_Sensitivity = val;
   	}
   }
}

