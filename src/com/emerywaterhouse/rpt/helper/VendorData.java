/**
 * File: VendorData.java
 * Description: Helper class that contains some vendor data information for the TopVendor 
 *    report.  Used here because of dynamic class loading issues when used as an inner class.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 01/08/2007
 * Last Update: $Id: VendorData.java,v 1.1 2007/01/08 16:27:42 jfisher Exp $
 * 
 * History 
 */
package com.emerywaterhouse.rpt.helper;

public class VendorData {      
   public int m_VndId;
   public String m_VndName;
   public double m_Sales;
   
   public VendorData()
   {
      m_VndId = 0;
      m_VndName = "";
      m_Sales = 0;
   }
}