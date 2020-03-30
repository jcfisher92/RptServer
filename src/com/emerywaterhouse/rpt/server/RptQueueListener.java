/**
 * File: RptQueueListener.java
 * Description: Listens for incomming messages on the report queue.  Pulls the BODs off
 *    of the queue and sends them to the report processor to be parsed and run.  This is the
 *    generic listener and does not listen for any special report.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 03/30/2005
 * Last Update: $Id: RptQueueListener.java,v 1.5 2008/10/30 16:55:03 jfisher Exp $
 * 
 * History
 */
package com.emerywaterhouse.rpt.server;

import java.io.IOException;

import org.apache.log4j.Logger;

import com.rabbitmq.client.AMQP.BasicProperties;
import com.rabbitmq.client.Channel;
import com.rabbitmq.client.DefaultConsumer;
import com.rabbitmq.client.Envelope;

public class RptQueueListener extends DefaultConsumer 
{
   private static int START_ID = 0;
   private RptMonitor m_Monitor; // The report que monitor class.
   protected String m_Name; // The name of the listener. Should equate to a topic.
   protected int m_Id; // The listener id number.

   //
   // Log4j logger
   private static Logger log = Logger.getLogger(RptQueueListener.class);

   /**
    * default constructor
    */
   public RptQueueListener() 
   {
      this(null, null, null);
   }

   public RptQueueListener(Channel channel) 
   {
      this(null, null, channel);
   }

   /**
    * Creates a listener with a monitor reference and a name.
    * 
    * @param monitor A reference to the ESBMonitor class
    * @param name The name of the to be given to the instance of the listener.
    */
   public RptQueueListener(RptMonitor monitor, String name, Channel channel)
   {
      super(channel);
      m_Monitor = monitor;
      m_Id = START_ID++;
      m_Name = name;
      log.info("starting listener: name = " + m_Name + " id = " + m_Id);
   }

   @Override
   public void handleDelivery(String consumerTag, Envelope env, BasicProperties props, byte[] body) throws IOException 
   {
      RptProcessor processor;
      String bod = null;
      boolean error = false;
      //the d-tag is important because you need it to ack the message.
      long deliveryTag = env.getDeliveryTag();
   
      try {
         if ( body != null ) {
            while ( bod == null && !error ) {
               if ( m_Monitor.getProcCount() < m_Monitor.getMaxRptCount() ) {
                  //make a bod string out of the message body
                  bod = new String(body);
                  
                  if ( bod != null && bod.length() > 0 ) {
                     m_Monitor.incRptCount();
                     processor = new RptProcessor(m_Monitor, bod);
                     processor.processBOD();

                     this.getChannel().basicAck(deliveryTag, false);
                  } 
                  else {
                     error = true;
                     log.error("missing bod xml from the bus");
                  }
               } 
               else {
                  try {
                     //
                     // Wait around and try again.
                     Thread.sleep(500);
                  }

                  catch (InterruptedException ex) {
                     break;
                  }
               }
            }
         }
      } 
      
      catch ( Exception e ) {
         //arguments are {dtag, multiple, repost}
         this.getChannel().basicNack(deliveryTag, false, true);
      }
   }
}
