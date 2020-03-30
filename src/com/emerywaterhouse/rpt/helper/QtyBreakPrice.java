/*
 * $Id: QtyBreakPrice.java,v 1.1 2008/08/13 00:56:05 jheric Exp $
 * 
 * Author: jacob
 * Date: Aug 12, 2008
 *
 * History:
 * 	$Log: QtyBreakPrice.java,v $
 * 	Revision 1.1  2008/08/13 00:56:05  jheric
 * 	Wrapped up Vendor Price Change report (raw and untested at this point because there is no delphi request screen).  Added qty break pricing objec and slot pricing object.
 *
 */
package com.emerywaterhouse.rpt.helper;

/**
 * @author jacob
 *
 */
public class QtyBreakPrice {
   private Integer m_MinQty = null;
   private Double m_Percent = null;

   /**
    * @param minQty - Integer minimum quantity for this price break
    * @param percent - Double percent break/discount for this quantity 
    */
   public QtyBreakPrice(Integer minQty, Double percent) {
      super();
      m_MinQty = minQty;
      m_Percent = percent;
   }
   
   /**
    * @return Integer - minimum quantity for this price break
    */
   public Integer getMinQty() {
      return m_MinQty;
   }
   
   /**
    * @param minQty - Integer minimum quantity for this price break 
    */
   public void setMinQty(Integer minQty) {
      m_MinQty = minQty;
   }
   
   /**
    * @return Double - percent break/discount for this quantity
    */
   public Double getPercent() {
      return m_Percent;
   }
   
   /**
    * @param percent - Double percent break/discount for this quantity 
    */
   public void setPercent(Double percent) {
      m_Percent = percent;
   }
   
}
