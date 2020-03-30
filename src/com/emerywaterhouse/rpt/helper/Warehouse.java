/**
 * File: Warehouse.java
 * Description: Simple container class to hold warehouse information.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 02/12/2009
 * Last Update: $Id: Warehouse.java,v 1.1 2009/02/12 19:46:29 jfisher Exp $
 * 
 * History
 *    $Log: Warehouse.java,v $
 *    Revision 1.1  2009/02/12 19:46:29  jfisher
 *    Inititial add
 *
 */
package com.emerywaterhouse.rpt.helper;

public class Warehouse
{
   public int emeryId;
   public String fascorId;
   public String accpacId;
   public String name;
   
   //
   // default constructor
   public Warehouse()
   {
      super();
   }
   
   /**
    * Constructor that initializes all the values.
    * 
    * @param emeryId
    * @param fascorId
    * @param accpacId
    * @param name
    */
   public Warehouse(int emeryId, String fascorId, String accpacId, String name)
   {
      this();
      
      this.emeryId = emeryId;
      this.fascorId = fascorId;
      this.accpacId = accpacId;
      this.name = name;
   }
   
   /**
    * Clean up;
    * @see java.lang.Object#finalize()
    */
   public void finalize() throws Throwable
   {
      fascorId = null;
      accpacId = null;
      name = null;
      
      super.finalize();
   }
}
