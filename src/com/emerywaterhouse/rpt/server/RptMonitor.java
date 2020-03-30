/**
 * File: RptMonitor.java
 * Description: Monitoring class for running reports.  It handles limiting the number of reports that can run,
 *    getting the status of reports, etc.  It also handle creating the connections to the message queue and starting
 *    them.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 03/29/2005
 * Last Update: $Id: RptMonitor.java,v 1.14 2012/10/17 20:56:08 jfisher Exp $
 *
 * History
 */
package com.emerywaterhouse.rpt.server;

import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeoutException;

import javax.naming.NamingException;

import com.rabbitmq.client.Channel;
import com.rabbitmq.client.Connection;
import com.rabbitmq.client.ConnectionFactory;

public class RptMonitor implements Runnable 
{
   private final static String rptQueue = "ha.RptReq";

   private Thread m_Thread;
   private int m_MaxRptCount;
   private int m_RptCount;
   private ArrayList<RptProcessor> m_RptProcs;
   private short m_Status;

   private Connection m_Cnx;
   private Channel m_Channel;

   /**
    * default constructor
    */
   public RptMonitor() 
   {
      super();

      m_Thread = new Thread(this, "RptMonitor");
      m_Thread.setDaemon(true);

      m_RptCount = 0;
      m_MaxRptCount = Integer.parseInt(System.getProperty("max.reports", "10"));

      //
      // Create the array list with a capacity equal to the max number of
      // reports.
      m_RptProcs = new ArrayList<RptProcessor>(m_MaxRptCount);
   }

   /**
    * Clean up anything we created.
    * 
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable 
   {
      m_Thread = null;
      
      if ( m_RptProcs != null ) {
         m_RptProcs.clear();
         m_RptProcs = null;
      }

      super.finalize();
   }

   /**
    * Adds a RptProcessor reference to the list of running reports.
    *
    * @param rptProc The running RptProcessor instance.
    */
   public synchronized void addRptProc(RptProcessor rptProc) 
   {
      if ( rptProc != null ) {
         synchronized (m_RptProcs) {
            m_RptProcs.add(rptProc);
         }
      }
   }

   /**
    * Closes the connections to the ActiveMQ message broker.
    */
   private void closeRMQConnection() 
   {
      try {
         m_Channel.close();
         m_Cnx.close();
      } 
      
      catch (IOException | TimeoutException e) {
         RptServer.log.error("[RptMonitor] Failed to close RMQ Connections.");
         RptServer.log.error("[RptMonitor]", e);
      }

      m_Cnx = null;
   }

   /**
    * Connects to the ActiveMQ message broker.
    *
    * @throws NamingException
    * @throws JMSException
    * @throws UnknownHostException
    */
   private void connectToRMQ() throws NamingException, UnknownHostException 
   {
      ConnectionFactory cnxFactory = null;
      String user = System.getProperty("msgbroker.user");
      String passwd = System.getProperty("msgbroker.passwd");
      String host = System.getProperty("msgbroker.host");
      int port = Integer.parseInt(System.getProperty("msgbroker.port"));
   
      try {
         cnxFactory = new ConnectionFactory();
         //see https://www.rabbitmq.com/api-guide.html#recovery
         cnxFactory.setAutomaticRecoveryEnabled(true);
         //userm, passwd, host, port are set on the factory, vhost is left as default
         cnxFactory.setUsername(user);
         cnxFactory.setPassword(passwd);
         cnxFactory.setHost(host);
         cnxFactory.setPort(port);
         m_Cnx = cnxFactory.newConnection();
         m_Channel = m_Cnx.createChannel();
         //this sets the maximum number of unacknowledged messages to 10 across the entire channel
         //it will act as a speed bump if we're not doing our acknowledgements correctly
         m_Channel.basicQos(10);
         //params are {queue_name, auto_ack, consumer} 
         m_Channel.basicConsume(rptQueue, false, getClientId(), new RptQueueListener(this, "reportreq", m_Channel));
      } 
      
      catch ( IOException e ) {
         RptServer.log.error("[RptMonitor]", e);
      } 
      
      catch ( TimeoutException e ) {
         RptServer.log.error("[RptMonitor]", e);
      }
   }

   /**
    * Decriments the report count var.
    */
   public synchronized void decRptCount() 
   {
      --m_RptCount;
   }

   /**
    * @return The id for the connection to the message broker. Based on host
    *         name and proc monitor name.
    *
    * @throws UnknownHostException
    */
   private String getClientId() throws UnknownHostException 
   {
      return String.format("%s.%s", InetAddress.getLocalHost().getHostName(), "RptSrv");
   }

   /**
    * Returns the maximum number of reports that can be running.
    *
    * @return The max report count.
    */
   public synchronized int getMaxRptCount() 
   {
      return m_MaxRptCount;
   }

   /**
    * Returns the current report count. This is the number of currently running
    * reports. This has to be limited because some reports take a lot of
    * resources and users run as many as they possibly can at once.
    *
    * @return The report count.
    */
   public synchronized int getProcCount() 
   {
      return m_RptCount;
   }

   /**
    * Gets the status information from each running report.
    *
    * @return HTML data that represents the status information from the
    *         currently running reports.
    */
   public synchronized String getRptStatus() 
   {
      Date date = null;
      DateFormat df = new SimpleDateFormat("MM/dd' 'HH:mm:ss");
      String maxRun = null;
      String curRun = null;
      long hour = 0;
      long min = 0;
      long sec = 0;

      StringBuffer buf = new StringBuffer(1024);
      RptStatus status = null;
      String row = "<td width=\"%s\">%s</td>\r\n";

      buf.append("<table style=\"font-family: Arial, Helvetica, sans-serif; ");
      buf.append("font-size: 12px; border-width: 0; border-style: none; width: 1000px\">\r\n");

      buf.append("<tr>\r\n");
      buf.append("<td width=\75px\"><b>TID</b></td>\r\n");
      buf.append("<td width=\"75px\"><b>IID</b></td>\r\n");
      buf.append("<td width=\"150px\"><b>Report Name</b></td>\r\n");
      buf.append("<td width=\"100px\"><b>User ID</b></td>\r\n");
      buf.append("<td width=\"100px\"><b>Started</b></td>\r\n");
      buf.append("<td width=\"100px\"><b>Running Time</b></td>\r\n");
      buf.append("<td width=\"100px\"><b>Max Run</b></td>\r\n");
      buf.append("<td width=\"300px\"><b>Current Action</b></td>\r\n");
      buf.append("<tr>\r\n");

      //
      // Format the data elements for display on the html page
      synchronized (m_RptProcs) {
         for (int i = 0; i < m_RptProcs.size(); i++) {
            status = m_RptProcs.get(i).getRptStatus();
            date = new Date(status.startTime);

            //
            // Convert to seconds and then break out the hours and minutes.
            // If this is not
            // done then the seconds will accumulate past 60. So will the
            // minutes.
            sec = status.runTime / RptServer.SECOND;

            if (sec >= 3600) {
               hour = sec / 3600;
               sec = sec - hour * 3600;
            }

            if (sec >= 60) {
               min = sec / 60;
               sec = sec - min * 60;
            }

            curRun = String.format("%02d:%02d:%02d", new Long(hour), new Long(min), new Long(sec));
            maxRun = Double.toString(status.maxRunTime / RptServer.HOUR) + " HRS";

            buf.append("<tr>");
            buf.append(String.format(row, "75px", Long.toString(status.threadId)));
            buf.append(String.format(row, "75px", Long.toString(status.internalId)));
            buf.append(String.format(row, "150px", status.rptName));
            buf.append(String.format(row, "100px", status.uid));
            buf.append(String.format(row, "100px", df.format(date)));
            buf.append(String.format(row, "100px", curRun));
            buf.append(String.format(row, "100px", maxRun));
            buf.append(String.format(row, "300px", status.currentAction));

            buf.append("</tr>\r\n");

            hour = 0;
            min = 0;
            sec = 0;
         }
      }

      buf.append("</table>");

      return buf.toString();
   }

   /**
    * Utility funtion for determining the time in hours, minutes or seconds.
    *
    * @param seconds The time in seconds.
    * @param timeType The type of time unit needed. These will be the RptServer.HOUR, MINUTE, SECOND.
    *
    * @return The time in the units specified or seconds if the unit is not one
    *         that matches.
    */
   @SuppressWarnings("unused")
   private long getTime(long seconds, int timeType) 
   {
      long time = 0;

      switch ( timeType ) {
         case RptServer.HOUR:
            if (seconds >= 3600)
               time = seconds / 3600;
            break;

         case RptServer.MINUTE:
            if (seconds >= 60)
               time = seconds / 60;
            break;

         default:
            time = seconds;
      }

      return time;
   }

   /**
    * Increments the process count var.
    */
   public synchronized void incRptCount() 
   {
      ++m_RptCount;
   }

   /**
    * Removes a RptProcessor object from the list of running reports.
    *
    * @param rptProc  The RptProcessor to remove.
    */
   public synchronized void remRptProc(RptProcessor rptProc) 
   {
      if ( rptProc != null ) {
         synchronized (m_RptProcs) {
            for ( int i = 0; i < m_RptProcs.size(); i++ ) {
               if ( rptProc.getId() == m_RptProcs.get(i).getId() )
                  m_RptProcs.remove(i);
            }
         }
      }
   }

   /**
    * Keeps tab on whether the server is started or stopped.
    * 
    * @see java.lang.Runnable#run()
    */
   @Override
   public void run() 
   {
      String msg = "[RptMonitor] monitor started";

      if ( Thread.currentThread() == m_Thread ) {
         RptServer.log.info(msg);

         //
         // Check the message queue. If there are messages, then start
         // processing the reports.
         while (m_Status != RptServer.STOPPED) {
            try {
               Thread.sleep(300);
            }

            catch ( Exception ex ) {
               if ( ex instanceof InterruptedException ) {
                  Thread.currentThread().interrupt();
                  break;
               }

               RptServer.log.error("[RptMonitor]", ex);
            }
         }
      } 
      else {
         msg = "[RptMonitor] the report processor should be started by calling the start method";
         RptServer.log.error(msg);
      }

      msg = "[RptMonitor] report processor stopped";
      RptServer.log.info(msg);
   }

   /**
    * Starts the report monitor and connects to the message queue.
    *
    * @throws Exception
    */
   public void start() throws Exception 
   {
      RptServer.log.info("[RptMonitor] starting up");
      RptServer.log.info("[RptMonitor] connecting to the message broker");
   
      connectToRMQ();
      
      RptServer.log.info("[RptMonitor] connected to the message broker");
      m_Status = RptServer.RUNNING;
      m_Thread.start();
   }

   /**
    * Stops the report processing and closes the queue connection.
    */
   public void stop() 
   {
      RptServer.log.info("[RptMonitor] report processor terminating");
      m_Status = RptServer.STOPPED;

      closeRMQConnection();
      RptServer.log.info("[RptMonitor] report processor stopped");
   }

   /**
    * Stops a report from running
    *
    * @param rptId The internal id number of the report processor.
    */
   public synchronized void stopReport(long rptId) 
   {
      RptProcessor proc = null;

      if ( rptId > -1 ) {
         synchronized (m_RptProcs) {
            for (int i = 0; i < m_RptProcs.size(); i++) {
               proc = m_RptProcs.get(i);

               if (proc.getId() == rptId) {
                  proc.setRptStatus(RptServer.ABORT);
                  RptServer.log.warn("[RptMonitor] report processor id: " + rptId + " was sent the stop command");
                  break;
               }
            }
         }
      }

      proc = null;
   }
}