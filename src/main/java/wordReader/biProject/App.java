package wordReader.biProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument; 
import org.apache.poi.hwpf.usermodel.Range; 
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.TableShape;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor; 
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import fr.opensagres.xdocreport.utils.StringEscapeUtils;

public class App 
{
	static boolean debug = false ;  // 要印出訊息的時候就設定為true , 平常是關起來false狀態
	
    public static void main( String[] args ) throws IOException
    {
    	// 歡迎訊息 並且會顯示 word取得路徑 以及excel產生路徑...等訊息
    	helloMsg();
        
        // 設一個while loop 
        // loop through file , read and get all word(.docx) under that filePath
        // Store all the data in dataPojosList
    	List<DataPojo> dataPojos = returnAllWordData();
        
    	
    	writeExcel(dataPojos);
    }
    
    public static void helloMsg() throws IOException {
        Properties props = new Properties();
        props.load(App.class.getClassLoader().getResourceAsStream("application.properties"));
        
        System.out.println( "Hi! " + props.getProperty("userName") );
        System.out.println( "word 存放路徑 : " + props.getProperty("wordsPath") );
        System.out.println( "excel 會存在: " + props.getProperty("writePath")  + "底下");
    }
    
    public static List<DataPojo> returnAllWordData() throws IOException{
    	
    	List<DataPojo> stackList = new ArrayList<>() ;

        Properties props = new Properties();
        props.load(App.class.getClassLoader().getResourceAsStream("application.properties"));
    	
        // word 存放路徑
    	String wordsPath = props.getProperty("wordsPath") ; 
    	
    	// 取得所有這路徑底下的 以.docx結尾之檔案
    	FileFilter fileFilter = new FileFilter() ;
    	File dir = new File(wordsPath) ;
    	File[] files = dir.listFiles(fileFilter);
    	if( files.length == 0 )
    		System.out.println("No .docx files under path : " + wordsPath);
    	else {
    		for( File aFile: files ) {
    			System.out.println("file : " + aFile.getName());
    			// 取得 word 裡面資料 並處理一些計算的部分
    			DataPojo readyDataPojo = calculateExtraFieldsData(  readWord2007Docx(wordsPath+aFile.getName()) ) ; 
    			stackList.add(readyDataPojo) ;
    		}	
    	}
    	
    	return stackList ;
    }
    
    /**
     * 寫Excel檔案到指定路徑
     * @param workbook
     * @param writePath
     * @throws IOException 
     */
    public static void writeExcel(List<DataPojo> dataPojos) throws IOException {
    	
        Properties props = new Properties();
        props.load(App.class.getClassLoader().getResourceAsStream("application.properties"));
    	
        // word 存放路徑
    	String writePath = props.getProperty("writePath") ; 
    	
    	try {
    		Workbook workbook = ExcelWriter.exportData(dataPojos) ; // POI會幫我們處理所有格式上所需
            FileOutputStream out=new FileOutputStream(writePath); 
    		workbook.write(out);
    		System.out.println("建立Excel成功");
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
    
    /**
     * 處理 從word上面獲得資訊之後, 補齊原本DataPojo沒有的欄位或者額外計算
     * @param dataPojo
     * @return dataPojo
     */
    public static DataPojo calculateExtraFieldsData(DataPojo dataPojo) {
    	
    	String year = Integer.toString(LocalDate.now().getYear()) ; 
    	
    	// 分開開始日期 與 時間
    	String startString = dataPojo.getStartTime() ;
    	String startMonthString = startString.substring(0, startString.indexOf('月')) ;
    	String startDateString = startString.substring(startString.indexOf('月') + 1, startString.indexOf('日')) ;
    	String startHourString = startString.substring(startString.indexOf('日') + 1, startString.indexOf('時')) ;
    	String startMinString = startString.substring(startString.indexOf('時') + 1, startString.indexOf('分')) ;
    	
    	dataPojo.setStartDay( year + "/" + startMonthString + "/" + startDateString);
    	dataPojo.setStartTime( startHourString + ":" + startMinString);
    	
    	// 分開結束日期 與 時間
    	String endString = dataPojo.getEndTime();
    	String endMonthString = endString.substring(0, endString.indexOf('月')) ;
    	String endDateString = endString.substring(endString.indexOf('月') + 1, endString.indexOf('日')) ;
    	String endHourString = endString.substring(endString.indexOf('日') + 1, endString.indexOf('時')) ;
    	String endMinString = endString.substring(endString.indexOf('時') + 1, endString.indexOf('分')) ;
    	
    	dataPojo.setEndDay( year + "/" + endMonthString + "/" + endDateString );
    	dataPojo.setEndTime( endHourString + ":" + endMinString );
    	
    	// 計算申請時數
    	// 時數 = 迄HR - 起HR
    	int endHour = Integer.valueOf(endHourString);
    	int startHour = Integer.valueOf(startHourString) ;
    	int endMin = Integer.valueOf(endMinString);
    	int startMin = Integer.valueOf(startMinString) ;
    
    	// 取得總共工時與分鐘
    	int applyHour = endHour - startHour ;
    	dataPojo.setApplyHour(Integer.toString(applyHour));
    	
    	// 計算承認工時
    	// 中午12~1點 以及 晚上6~7點是不算時數的(吃飯時間)
    	Time startTime = new Time(startHour, startMin);
    	Time endTime = new Time(endHour, endMin);
    	Time admitTime = Time.getTotalHourNMins(startTime, endTime);
    	if( admitTime.minutes >= 30 ) {
    		// 超過30分鐘 直接進位多一小時
    		admitTime.hours = admitTime.hours + 1 ;
    	}
    	
    	dataPojo.setAdmitTime(Integer.toString( admitTime.hours ));
    	
    	// 設定專案名稱 
    	// 目前測試階段都設定為合庫
    	dataPojo.setProjectName("合庫");
    	
    	if( debug ) {
    		System.out.println(dataPojo.getStartDay()) ;
    		System.out.println(dataPojo.getStartTime()) ;
    		System.out.println(dataPojo.getEndDay()) ;
    		System.out.println(dataPojo.getEndTime()) ;
    		System.out.println(dataPojo.getApplyHour()) ;
    	}
    
    	return dataPojo ;
    }
    
    /**
     * 取得word裡面我們要的資訊, 存到並回傳 DataPojo
     * @param filePathDocx
     * @return
     */
    public static DataPojo readWord2007Docx(String filePathDocx){ 

    	DataPojo dataPojo = null ; 
   	 	try {
   	 	   File file = new File(filePathDocx);
   	 	   FileInputStream fis = new FileInputStream(file.getAbsolutePath());
   	 	   XWPFDocument document = new XWPFDocument(fis);
   	 	   XWPFWordExtractor extractor = new XWPFWordExtractor(document);
   	 	   
	 	   String bodyString = extractor.getText() ;
	 	   String xmlString = document.getDocument().toString() ;
   	 	   if( debug ) {
   	 		   System.out.println("Date : " + getDate(bodyString));
   	 		   System.out.println("Name : " + getStaffName(bodyString));
   	 		   System.out.println("Reason : " + getLateReason(xmlString));
   	 		   System.out.println("Apartment : " + getApartment(bodyString));
   	 		   System.out.println("Start Time : " + getStartTime(bodyString));
   	 		   System.out.println("End Time : " + getEndTime(bodyString));
   	 		   System.out.println("Rest or Get money : " + restOrMoney(bodyString));
   	 	   }
   	 	   
   	 	   // Write data to dataPojo 
   	 	   dataPojo = new DataPojo() ;
   	 	   dataPojo.setDate( getDate(bodyString) );
   	 	   dataPojo.setName( getStaffName(bodyString) );
   	 	   dataPojo.setReason( getLateReason(xmlString) );
   	 	   dataPojo.setApartment( getApartment(bodyString));
   	 	   dataPojo.setStartTime( getStartTime(bodyString) );
   	 	   dataPojo.setEndTime( getEndTime(bodyString ) );
   	 	   dataPojo.setRestOrMoney(restOrMoney(bodyString));
   	 	   
           fis.close();
		} catch(Exception ex) {
		    ex.printStackTrace();
		} 
   	 	
   	 	return dataPojo ; 
    }
    
    // 去除String裡面空白換行等字元
    private static String cleanString(String unCleanString) { 
    	return unCleanString.trim().replaceAll("\r\n|\r|\n|\t|\f|\b", "").replaceAll("\\s+","") ;
    }
    
    // 取得年月日
    private static String getDate(String document) {
    	
    	int applyDateIndex = document.indexOf("申請日期:") ; 
    	int staffNameIndex = document.indexOf("員工姓名") ; 
    	int offset = "申請日期:".length(); 
    	
    	String originalDateString = cleanString(document.substring(applyDateIndex+offset, staffNameIndex)) ; 
    	int yearIndex = originalDateString.indexOf("年") ; 
    	int monthIndex = originalDateString.indexOf("月");
    	int dateIndex = originalDateString.indexOf("日") ;
    	
    	String monthString = originalDateString.substring(yearIndex+1, monthIndex) ;
    	String dateString = originalDateString.substring(monthIndex+1,dateIndex) ; 
    	
    	return monthString + "/" + dateString ; 
    }
    
    // 取得員工姓名
    private static String getStaffName( String document) {
    	int staffNameIndex = document.indexOf("員工姓名") ; 
    	int apartmentIndex = document.indexOf("部門") ; 
    	int offset = "員工姓名".length() ;
    	
    	return cleanString(document.substring(staffNameIndex+offset, apartmentIndex)) ;
    }
    
    // 取得部門
    private static String getApartment(String document) {
    	int apartmentIndex = document.indexOf("部門") ; 
    	int expectTimeIndex = document.indexOf("預計") ; 
    	int offset = "部門".length() ;
    	
    	return cleanString(document.substring(apartmentIndex+offset, expectTimeIndex));
    }
    
    // 取得起始時間
    private static String getStartTime(String document) {
    	int startTimeIndex = document.indexOf("時間起") ; 
    	int endTimeIndex = document.indexOf("時間迄") ;
    	int offset = "時間起".length() ;
    	
    	return cleanString(document.substring(startTimeIndex+offset, endTimeIndex));
    }
    
    // 取得結束時間
    private static String getEndTime(String document) {
    	int endTimeIndex = document.indexOf("時間迄") ;
    	int reasonIndex = document.indexOf("加班事由") ;
    	int offset = "時間迄".length() ;
    	
    	return cleanString(document.substring(endTimeIndex+offset, reasonIndex));
    }
 
    // 取得加班理由
    private static String getLateReason(String xmlString) {
       String tmp_String = "" ; // for concat 用 
    	
 	   String textBoxEndString = "</v:textbox>" ;
 	   int start = xmlString.indexOf("<v:textbox>");
 	   int end = xmlString.indexOf("</v:textbox>");
 	   int endMatchLength = textBoxEndString.length() ; 
 	   
 	   String lateReasonString = xmlString.substring(start, end+endMatchLength); // 取得xml裡面textBox的部分
 	   String patternStr = "<w:t>.*</w:t>"; // regex
 	   Pattern pattern = Pattern.compile(patternStr) ;
 	   Matcher matcher = pattern.matcher(lateReasonString) ; 
 	   
 	   int startSize = "<w:t>".length() ; 
 	   int endSize = "</w:t>".length() ; 
 	   while( matcher.find() ) { // 找到所有符合之項目
	        String targetString = matcher.group() ; 
	        tmp_String = tmp_String + targetString.substring(startSize, targetString.length()-endSize) ; 
 	   }
 	 
 	   return cleanString(tmp_String);
    }
    
    // 補修或者加班費
    private static boolean restOrMoney(String document) {
    	int methodIndex = document.indexOf("使用方式") ;
    	int actualTimeIndex = document.indexOf("實際工作時數") ;
    	int offset = "使用方式".length() ;
    	
    	String resultString = cleanString(document.substring(methodIndex+offset, actualTimeIndex));
    	
    	if( resultString.indexOf("補休") == 0 ) { // 代表補修被勾選,但string顯示不出來
    		return true ; 
    	}
    	
    	return false ;
    }
    
}
