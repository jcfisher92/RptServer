package com.emerywaterhouse.rpt.alert;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DataSender;
import com.emerywaterhouse.websvc.Param;

public class Screamer  extends Report
{
   private String m_RptName;
   public Session session;
   
   public Screamer() 
   {
      super();
   }
   
   //
   // build the error message
   private String buildMessage()
   {      
      StringBuffer html = new StringBuffer();
      
      html.append("<!DOCTYPE html>");
      html.append("<html>");
      html.append("<body>");      
      html.append("<div style=\"font-family: Arial, Helvetica, sans-serif; font-size: 12px; text-align: left; vertical-align: top\">");
      html.append("<h1>Oops! Something went wrong with your report.</h1>");
      html.append("<br>");
      html.append("<br>");
      html.append("<h2>Argh, where is my report? HHHEEELLLPPPPP</h2>");
      html.append("<img src=\"http://www.emeryonline.com/shared/images/catalog/email_img/screamer.png\" alt=\"Argh, where is my report?\" width=\"200\" height=\"200\">");      
      html.append("<br>");
      html.append("<br>");
      html.append("Contact the computer operations team (Thom, Jeff K, Debbie and Larry) about the problem");
      html.append("<br>");
      html.append("cot@emeryonline.com");
      html.append("</div>");
      html.append("</body>");
      html.append("</html>");
      
      return html.toString();
   }
   
   /**
    * Creates the report file.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean result = false;
      m_Status = RptServer.RUNNING;

      try {
         result = sendEmailAlert();         
      }

      catch ( Exception ex ) {
         log.fatal("[Screamer]", ex);
      }

      finally {
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return result;
   }

   /**
    * Creates the alert email and sends it to the distribution list.
    * @return boolean.  True if the email was sent, false if not.
    * @throws IOException 
    * @throws FileNotFoundException
    */
   public boolean sendEmailAlert() throws IOException
   {
      boolean sent = false;      
      String[] distList = m_RptProc.getDistList();
      String from = "noreply@emeryonline.com";
      final Properties props = DataSender.loadSmtpProps();
      InternetAddress[] addrList = new InternetAddress[distList.length];
      Message message = null;
      
      // Get the Session object.
      session = Session.getInstance(props, new Authenticator());
      
      try {
         for ( int i = 0; i < distList.length; i++)
            addrList[i] = new InternetAddress(distList[i]);
         
         message = new MimeMessage(session);         
         message.setFrom(new InternetAddress(from));
         message.setRecipients(Message.RecipientType.TO, addrList);
         message.setSubject(String.format("%s Report Submission Error", m_RptName));         
         message.setContent(buildMessage(), "text/html");
         
         Transport.send(message);
         sent = true;
      } 
      
      catch ( MessagingException ex ) {
         ex.printStackTrace();         
      }
      
      finally {
         message = null;
         session = null;
         
         for ( int i = 0; i < addrList.length; i++)
            addrList[i] = null;
         
         addrList = null;
      }
      
      return sent;
   }
   
   /**
    * Sets the parameters for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    * Only need to get the submitted report name so it can be used in the alert email.
    */
   public void setParams(ArrayList<Param> params)
   {
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("reportname") ) {
            m_RptName = param.value;
            break;
         }
      }
   }
   
   public class Authenticator extends javax.mail.Authenticator 
   {      
      public Authenticator()
      {
         super();
      }
      
      public PasswordAuthentication getPasswordAuthentication() {
         return new PasswordAuthentication(System.getProperty("mail.smtp.user"), System.getProperty("mail.smtp.password"));
      }
   }
}
