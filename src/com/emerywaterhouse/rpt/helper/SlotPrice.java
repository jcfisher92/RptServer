/*
 * $Id: SlotPrice.java,v 1.1 2008/08/13 00:56:05 jheric Exp $
 * Description:  Represents a slot price, a minimum quantity and price used for commodity item slot pricing.  This is used 
 *   in the Vendor Price Change report (see com.emerywaterhouse.rpt.spreadsheet.VndPriceChange).
 *   
 * Author: jacob
 * Date: Aug 12, 2008
 *
 * History:
 * 	$Log: SlotPrice.java,v $
 * 	Revision 1.1  2008/08/13 00:56:05  jheric
 * 	Wrapped up Vendor Price Change report (raw and untested at this point because there is no delphi request screen).  Added qty break pricing objec and slot pricing object.
 *
 */
package com.emerywaterhouse.rpt.helper;

/**
 * @author jacob
 *
 */
public class SlotPrice {
   private Double m_Price = null;
   private Integer m_MinQty = null;
   
   /**
    * @param price - Double price for this slot
    * @param minQty - Integer minimum quantity for this slot
    */
   public SlotPrice(Double price, Integer minQty) {
      super();
      m_Price = price;
      m_MinQty = minQty;
   }

   /**
    * @return Double - price for given slot.
    */
   public Double getPrice() {
      return m_Price;
   }
   
   /**
    * @param price - Double price for this slot
    */
   public void setPrice(Double price) {
      m_Price = price;
   }
   
   /**
    * @return Integer - Minimum quantity for this slot
    */
   public Integer getMinQty() {
      return m_MinQty;
   }
   
   /**
    * @param minQty - Integer minimum quantity for this slot
    */
   public void setMinQty(Integer minQty) {
      m_MinQty = minQty;
   }
   
   

}
