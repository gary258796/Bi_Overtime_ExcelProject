package wordReader.biProject.action.inMainAction.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import wordReader.biProject.model.DataPojo;
import wordReader.biProject.util.PropsHandler;
import wordReader.biProject.util.Time;

public class WordReader {
	
	static boolean debug = false ;  // 要印出訊息的時候就設定為true , 平常是關起來false狀態
	
    /**
     * 取得word裡面我們要的資訊, 計算額外欄位資訓, 存到並回傳 DataPojo
     * @param filePathDocx
     * @return
     * @throws IOException 
     */
    @SuppressWarnings("resource")
	public static DataPojo readWord2007Docx(String filePathDocx) throws IOException{ 

    	DataPojo dataPojo = null ; 
    	
   	 	try {
   	       // 取得該word的內容 以及 xmlString
   	 	   File file = new File(filePathDocx);
   	 	   FileInputStream fis = new FileInputStream(file.getAbsolutePath());
   	 	   XWPFDocument document = new XWPFDocument(fis);
   	 	   XWPFWordExtractor extractor = new XWPFWordExtractor(document);
   	 	   
	 	   String bodyString = extractor.getText() ;
	 	   	 	   
   	 	   if( debug ) {
   	 		   System.out.println("bodyString " + bodyString + "\n");
   	 		   System.out.println("Date : " + getApplyDate(bodyString));
   	 		   System.out.println("Apartment : " + getApartment(bodyString));
   	 		   System.out.println("Name : " + getStaffName(bodyString));
   	 		   System.out.println("StartDate : " + getActualDate(bodyString));
   	 		   System.out.println("StartTime : " + getStartTime(bodyString));
   	 		   System.out.println("EndTime : " + getEndTime(bodyString));
   	 		   System.out.println("ProjectName : " + getProjectName(bodyString));
   	 		   System.out.println("Late Reason : " + getLateReason(bodyString));
   	 		   System.out.println("Rest or Money : " + restOrMoney(bodyString));
   	 		   System.out.println("Extra Msg : " + getExtraMsg(bodyString));
   	 		   System.out.println("Has Photo : " +  isImageInOrNot( document.getDocument().toString()) ) ;
   	 	   }

   	 	   // Write data to dataPojo 
   	 	   dataPojo = new DataPojo() ;
   	 	   dataPojo.setDate( getApplyDate(bodyString) );
   	 	   dataPojo.setApartment( getApartment(bodyString));
   	 	   dataPojo.setName( getStaffName(bodyString) );
   	 	   dataPojo.setStartDay(getActualDate(bodyString));
   	 	   dataPojo.setStartTime(getStartTime(bodyString));
   	 	   dataPojo.setEndTime(getEndTime(bodyString));
   	 	   dataPojo.setProjectName(getProjectName(bodyString));
   	 	   dataPojo.setReason( getLateReason(bodyString) );
   	 	   dataPojo.setRestOrMoney(restOrMoney(bodyString));
   	 	   dataPojo.setExtraMsg(getExtraMsg(bodyString));
   	 	   dataPojo.setHasPhoto( isImageInOrNot( document.getDocument().toString() ) ) ; 
   	 	   dataPojo.setActualStartTime("") ;
   	 	   dataPojo.setActualEndTime(""); // 實際下班時間
   	 	   dataPojo.setDifferTotalTime(0);
   	 	   dataPojo.setSunday(false) ;
   	 	   dataPojo.setMissContent("");
   	 	   
           fis.close();
		} catch(Exception ex) {
			
			// 有問題檔案 丟到設定路徑底下
	        String destination = PropsHandler.getter("errorWordsPath") ;
        	String fileName = filePathDocx.substring(filePathDocx.lastIndexOf("/"), filePathDocx.length());
        	
			FileChannel in = new FileInputStream( filePathDocx ).getChannel();
        	FileChannel out = new FileOutputStream( destination+fileName ).getChannel();
        	out.transferFrom( in, 0, in.size() );
        	in.close();
        	out.close();

        	System.out.println(fileName + " 有問題,已加入到指定error word路徑！");
			dataPojo = null ;
		} 
   	 	
  
   	 	return calculateExtraFieldsData(dataPojo) ; 
    }
    
    /**
     * 處理 從word上面獲得資訊之後, 補齊原本DataPojo沒有的欄位或者額外計算
     * @param dataPojo
     * @return dataPojo
     */
    public static DataPojo calculateExtraFieldsData(DataPojo dataPojo) {
    	
    	if( dataPojo != null ) {
        	
        	// 設定起始 小時 與分鐘 方便後面計算
        	String startString = dataPojo.getStartTime() ;
        	String startHourString = startString.substring(0, startString.indexOf(':')) ;
        	String startMinString = startString.substring(startString.indexOf(':') + 1, startString.length()) ;
        	dataPojo.setEndDay( dataPojo.getStartDay() );
        	dataPojo.setStartHour(Integer.parseInt(startHourString));
        	dataPojo.setStartMin(Integer.parseInt(startMinString));
        	
        	// 分開結束日期 與 時間
        	String endString = dataPojo.getEndTime();
        	String endHourString = endString.substring( 0, endString.indexOf(':')) ;
        	String endMinString = endString.substring(endString.indexOf(':') + 1, endString.length()) ;
        	dataPojo.setEndHour( Integer.parseInt(endHourString));
        	dataPojo.setEndMin( Integer.parseInt(endMinString));

        	int endHour = Integer.parseInt(endHourString);
        	int startHour = Integer.parseInt(startHourString) ;
        	int endMin = Integer.parseInt(endMinString);
        	int startMin = Integer.parseInt(startMinString) ;
			Time startTime = new Time(startHour, startMin);
			Time endTime = new Time(endHour, endMin);
        	// 申請時數 = 總共工時(結束時間 - 開始時間)
			Time applyHour = Time.diffTime(startTime, endTime);
        	dataPojo.setApplyHour( applyHour.getHours() + "時:" +  applyHour.getMinutes() + "分");
        	
        	// 計算承認工時
        	// 中午12~1點 以及 晚上6~7點是不算時數的(吃飯時間)
			Time admitTime = Time.getTotalHourNMins(startTime, endTime);
        	if( admitTime.getMinutes() >= 30 ) {
        		// 超過30分鐘 直接進位多一小時
        		admitTime.setHours( admitTime.getHours() + 1 );
        	}
        	dataPojo.setAdmitTime(Integer.toString( admitTime.getHours() ));
 
        	
        	return dataPojo ;
    	}
    
    	return null ; 
    }

   
    // 去除String裡面空白換行等字元
    private static String cleanString(String unCleanString) { 
    	return unCleanString.trim().replaceAll("\r\n|\r|\n|\t|\f|\b", "") ;
    }
    

    // 申請日期
    private static String getApplyDate(String document) {
    	
    	String startString = "申請日期" ; 
    	String endString = "部門名稱" ;  
    		
    	String dateString =  getFiledCommon(document, startString, endString);

    	String yearString = dateString.substring(0, dateString.indexOf('/')) ; 
    	String monthString = dateString.substring( dateString.indexOf('/') + 1, nthOccurrence(dateString, "/", 2) ) ;
    	String dayString = dateString.substring( dateString.lastIndexOf('/') + 1, dateString.length() ) ;

    	if( Integer.parseInt(monthString) < 10 ) {
    		monthString = "0" + monthString ; 
    	}
    	
    	if( Integer.parseInt(dayString) < 10 ) {
    		dayString = "0" + dayString ; 
    	}
    	
    	return yearString + "/" + monthString + "/" + dayString ; 
    }
    
    // 部門名稱
    
    // 取得部門
    private static String getApartment(String document) {
    	
    	String startString = "部門名稱" ; 
    	String endString = "員工姓名";
    	
    	return getFiledCommon(document, startString, endString);
    }
    
    // 取得員工姓名
    private static String getStaffName( String document) {
    	String startString = "員工姓名(中文)" ; 
    	String endString = "實際加班日期";
    	
    	return getFiledCommon(document, startString, endString);
    }
    
    // 實際加班日期
    
    // 取得實際加班日期
    private static String getActualDate(String document) {
    	String startString = "實際加班日期" ; 
    	String endString = "實際加班時間"; 
    		
    	String dateString =  getFiledCommon(document, startString, endString);
    	
    	String yearString = dateString.substring(0, dateString.indexOf('/')) ; 
    	String monthString = dateString.substring( dateString.indexOf('/') + 1, nthOccurrence(dateString, "/", 2) ) ;
    	String dayString = dateString.substring( dateString.lastIndexOf('/')+ 1, dateString.length() ) ;
    	
    	if( Integer.parseInt(monthString) < 10 ) {
    		monthString = "0" + monthString ; 
    	}
    	
    	if( Integer.parseInt(dayString) < 10 ) {
    		dayString = "0" + dayString ; 
    	}
    	
    	return yearString + "/" + monthString + "/" + dayString ; 
    }
    
    // 實際 開始 加班時間
    
    // 取得實際加班時間(開始)
    private static String getStartTime(String document) {
    	String startString = "(起)" ; 
    	String endString = "(迄)";
    	
    	String hourAndMinuteString =  getFiledCommon(document, startString, endString);
    	int index = hourAndMinuteString.indexOf("時") ; 

    	String hourString = cleanString(hourAndMinuteString.substring(0, index)) ;
    	String minString = cleanString(hourAndMinuteString.substring(index+1, hourAndMinuteString.indexOf("分") )) ;
    	
    	return hourString + ":" + minString  ;
    }
    
    // 實際結束加班時間
    
    // 取得實際加班時間(結束)
    private static String getEndTime(String document) {
    	String startString = "(迄)" ; 
    	String endString = "專案名稱" ;
    	
    	String hourAndMinuteString =  getFiledCommon(document, startString, endString);
    	int index = hourAndMinuteString.indexOf("時") ; 

    	String hourString = cleanString(hourAndMinuteString.substring(0, index)) ;
    	String minString = cleanString(hourAndMinuteString.substring(index+1, hourAndMinuteString.indexOf("分") )) ;
    	
    	return hourString + ":" + minString  ;
    }
    
    // 專案名稱
    
    // 取得 專案名稱
    private static String getProjectName(String document) {
    	String startString = "(請填寫完整)" ; 
    	String endString = "加班事由";
    	
    	return getFiledCommon(document, startString, endString);
    }
    
    // 加班理由
    
    // 取得 加班事由
    private static String getLateReason(String document) {
    	String startString = "加班事由" ; 
    	String endString = "使用方式";
    	
    	return getFiledCommon(document, startString, endString);
    }
    
    // 使用方式
    
    // 使用方式
    private static String restOrMoney(String document) {
    	String startString = "使用方式" ; 
    	String endString = "備註";
    	
    	return getFiledCommon(document, startString, endString);
    }
    
    // 備註
    
    // 取得 備註
    private static String getExtraMsg(String document) {
    	String startString = "備註" ; 
    	String endString = "注意事項";
    	
    	return getFiledCommon(document, startString, endString);
    }
    
    // 有無加入截圖
    
    // Ｗord 有無上傳圖片
    private static boolean isImageInOrNot( String xmlString ) {
    	
    	if( xmlString.contains("圖片")) 
    		return true ; 
    	
    	return false ; 
    	
    }
    
    // 共用的取得欄位部分
    
    // 取得欄位的共用部分
    private static String getFiledCommon(String document, String startString, String endString) {
    	
    	int startIndex = document.indexOf(startString) ; 
    	int endIndex = document.indexOf(endString) ; 
    	int offset = startString.length() ;
    	
    	return cleanString(document.substring(startIndex+offset, endIndex));	
    }
    
    // 取得字串裡面要尋找的第n個子字串Index
    
    // 取得在一個String裡面 符合條件的第n個 Index
    private static int nthOccurrence(String str1, String str2, int n) {
    	 
        String tempStr = str1;
        int tempIndex = -1;
        int finalIndex = 0;
        for(int occurrence = 0; occurrence < n ; ++occurrence){
            tempIndex = tempStr.indexOf(str2);
            if(tempIndex==-1){
                finalIndex = 0;
                break;
            }
            tempStr = tempStr.substring(++tempIndex);
            finalIndex+=tempIndex;
        }
        return --finalIndex;
    }
    

}
