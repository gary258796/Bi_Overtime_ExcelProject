package wordReader.biProject;

import java.io.IOException;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Flags.Flag;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import wordReader.biProject.util.EmailFinder;
import wordReader.biProject.util.PropsHandler;

public class OutlookSender {
	
	/**
	 * 
	 * @param name 要寄送對象的名字
	 * @param lateDate 加班的日期
	 * @throws IOException
	 */
	public static void sendMail(String name, String lateDate) throws IOException {
	
		// Outlook 相關參數設定
	    Properties props = new Properties();
	    props.put("mail.smtp.auth", "true");
	    props.put("mail.smtp.starttls.enable", "true");
	    props.put("mail.smtp.host", "mail.softbi.com");
	    props.put("mail.smtp.port", "587");
	    
	    final String mailAccount = PropsHandler.getter("emailAccount") ;
	    final String passWord = PropsHandler.getter("emailPassWord") ;

		
	    // 連上Outlook
		Session session = Session.getInstance( props,  new javax.mail.Authenticator() {
	        @Override
	        protected PasswordAuthentication getPasswordAuthentication() {
	            return new PasswordAuthentication(mailAccount, passWord);
	        }
		}) ;

		
	    try {

	    	// 定義信件內容 標題...等
	        MimeMessage message = new MimeMessage(session);
	        message.setFrom(new InternetAddress(mailAccount));
	        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(EmailFinder.returnEmail(name)));
	        message.setSubject("【通知】請補齊加班單資料~", "UTF-8");
	        
	        MimeMultipart multipart = new MimeMultipart("related") ;
	        BodyPart messageBodyPart = new MimeBodyPart() ;
	        String htmlText = 
	        		"<div style=\" font-family:Microsoft JhengHei, Helvetica, sans-serif; font-size:16px; color:#5B9BD5; \">Hi " + name + " ,<br>你申請的加班單(加班日期：" + lateDate + ")，這封沒有附上打卡時間的截圖，<br>請補上截圖內容，再將加班申請單寄給管理部，謝謝~</div>"
	        		+ "<br>" + 
	        		"<div><font size=\"4.5\" color=#AEAAAA face=\"Comic Sans MS\">Thanks and Regards.</font></div>" +
	        		"<div><font size=\"4.5\" color=#AEAAAA face=\"Comic Sans MS\"> Phina.deng</font></div>" +
	        		"<div><font size=\"4.5\"color=#AEAAAA face=\"Comic Sans MS\">---------------------</font></div><br>" + 
	        		"<img src=\"cid:image\"><div style=\" font-family:Microsoft JhengHei, Helvetica, sans-serif; font-size:14px; color:#AEAAAA; \">SoftBI Corp. Ltd.</div>"
	        		+"\n" +
	        		"<div style=\" font-family:Microsoft JhengHei, Helvetica, sans-serif; font-size:13px; color:#AEAAAA; \">商智資訊股份有限公司</div>"
	        		+"\n" +
	        		"<div style=\" font-family:Calibri, Helvetica, sans-serif; font-size:15px; color:#AEAAAA; \">SoftBI Corporation Limited<br>Tel :  886-2-23785588 ext.167<br>Fax:  886-2-23785583</div>";

	        messageBodyPart.setContent(htmlText, "text/html;charset=UTF-8");
	        multipart.addBodyPart(messageBodyPart);
	        
	        messageBodyPart = new MimeBodyPart() ; 
	        DataSource fdsDataSource  = new FileDataSource(PropsHandler.getPropertiesPath("BiImagePath")) ; 
	        messageBodyPart.setDataHandler(new DataHandler(fdsDataSource));
	        messageBodyPart.setHeader("Content-ID", "<image>");
	        multipart.addBodyPart(messageBodyPart);
	        
	        message.setContent(multipart);

	        // 發送
	        Transport.send(message);

	        // 將發送的訊息 存到寄件備份
	        Store store = session.getStore("imap");
	        store.connect("mail.softbi.com", 143 , mailAccount, passWord);
	        Folder folder = store.getFolder("寄件備份");
	        folder.open(Folder.READ_WRITE);  
	        message.setFlag(Flag.SEEN, true);  
	        folder.appendMessages(new Message[] {message});  
	        store.close();

	    } catch (MessagingException e) {
	        throw new RuntimeException(e);
	    }
			
	}
	
	
	
}
