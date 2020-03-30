/*
 * $Id: ReasonDispCode.java,v 1.2 2008/10/30 15:58:22 jfisher Exp $
 * Description:  Represents a reason and disposition code (and warehouse).  This is currently used by reports (see ItemReturnAnalysis).
 * 
 * Author: jacob
 * Date: Aug 7, 2008
 *
 * History:
 * 	$Log: ReasonDispCode.java,v $
 * 	Revision 1.2  2008/10/30 15:58:22  jfisher
 * 	Fixed some warnings
 *
 * 	Revision 1.1  2008/08/08 00:53:58  jheric
 * 	Break reasondispcode out of itemreturanalysis and move it under helper.
 *
 */
package com.emerywaterhouse.rpt.helper;

import com.emerywaterhouse.fascor.Facility;


public class ReasonDispCode {
   
   private Facility facility; 
   private String reason;
   private String disp;
   
   /**
    * @param f The facility
    * @param reasonCode - String the credit reason code.
    * @param dispCode - String the credit disposition code.
    */
   public ReasonDispCode(Facility f, String reasonCode, String dispCode){
      this.setFacility(f);
      this.setReason(reason = reasonCode != null ? reasonCode : "");
      this.setDisp(dispCode != null ? dispCode : "");
   }

   /**
    * @return - String credit reason code.
    */
   public String getReason() {
      return reason;
   }

   /**
    * @param reason - String credit disposition code.
    */
   public void setReason(String reason) {
      this.reason = reason;
   }

   /**
    * @return - String credit disposition code.
    */
   public String getDisp() {
      return disp;
   }

   /**
    * @param disp - String credit disposition code. 
    */
   public void setDisp(String disp) {
      this.disp = disp;
   }

   /**
    * @return Facility - facility object.
    */
   public Facility getFacility() {
      return facility;
   }

   /**
    * @param facility - Facility object.
    */
   public void setFacility(Facility facility) {
      this.facility = facility;
   }

}
