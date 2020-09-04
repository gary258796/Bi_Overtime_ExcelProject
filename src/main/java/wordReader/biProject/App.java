package wordReader.biProject;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.configuration.ConfigurationException;
import org.apache.poi.ss.usermodel.Workbook;

import wordReader.biProject.cusError.ExcelException;
import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.cusError.WordFileException;
import wordReader.biProject.fileNameFilter.WordsFileFilter;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.model.PinkPojo;
import wordReader.biProject.util.PropsHandler;
import wordReader.biProject.util.Time;

public class App 
{
	
	private List<DataPojo> dataPojos ; 
	private List<PinkPojo> pinkPojos ; 
	private HandlePink handlePink = new HandlePink() ;
			
    public void main() throws IOException, ParseException, StopProgramException, WordFileException, ConfigurationException
    {
    	// 歡迎訊息 並且會顯示 word取得路徑 以及excel產生路徑...等訊息
    	helloMsg();
      
    	changeProperties() ; 
    	
    	try {
    		// 取得加班單資料,並且計算完加班單裡面沒有的欄位資訊
    		// 無法成功獲取的word會放入到指定路徑底下
    		dataPojos = returnAllWordData();

        	// 取得震旦雲原檔裡面,粉紅色row的每一筆資料(上/下班欄位至少有一個是有值的)
    		pinkPojos = handlePink.handlePinkExcel();
		} catch (IOException e) {
			throw new IOException( e.getMessage() ) ;
		} catch (StopProgramException e) {
			throw new StopProgramException( e.getMessage() ) ;
		} catch (ExcelException e) {	// 路徑底下沒有 .xls
			System.out.println( e.getMessage() );
		} finally {
	    	if( CollectionUtils.isEmpty(pinkPojos) || pinkPojos.size() == 0 )
	    		System.out.println("PinkPojos 為null或者大小為0,可能是因為這個excel粉紅色欄位index不是59.");
		}

    	// 將準備好的資料寫到 excel , Excel Writer 負責 呈現的部分
    	finalProcess(pinkPojos, dataPojos);

    	// 將DataPojo 裡面 , hasPhoto 為false的資料 送outlook mail給該位同仁
    	sendMail(dataPojos) ;

		System.out.println("程式結束.");
    }
    
    
    private void changeProperties() throws IOException, ConfigurationException {
    	
        BufferedReader consoleInput=new BufferedReader(new InputStreamReader(System.in));
    	
    	boolean change = false ; 
        char yesNo ; 
    	whileloop:
    	while( true) {
    		System.out.println("請問有需要更改以上任一資訊嗎？(Y/N)");
    		yesNo = consoleInput.readLine().charAt(0);

        	if( yesNo == 'Y' ) {
        		change = true ; 
        		break whileloop ;
        	}
        	else if( yesNo == 'N' ) {
    			break whileloop ;
    		}
        	else 
        		System.out.println("\n 你是故意的還是在靠北我？");
    	}
    	
        int chooseNum ; 
        changeloop:
        while( change ) {
        	System.out.println("需要修改哪個？");
        	System.out.println("1. 信箱&密碼 2. 加班單Word存放路徑 3. 產出Excel路徑 "
        			+ "4. 無法處理Word存放路徑 5. 通訊錄Excel路徑 6. 粉紅顏色(你應該不會用到,是給專業的屎用的) "
        			+ "7. 查看現在狀態 8. 結束");
        
        	chooseNum = Integer.parseInt(consoleInput.readLine()) ; 
        	String inputString = "" ;
        	switch (chooseNum) {

				case 1:
					System.out.print("請輸入信箱(空白並且按下Enter取消更改) : ");
					String emailString = consoleInput.readLine();
					System.out.print("請輸入密碼(空白並且按下Enter取消更改) : ");
					String passWordString = consoleInput.readLine();
					PropsHandler.setter("emailAccount", emailString);
					PropsHandler.setter("emailPassWord", passWordString);
					break;
				case 2:
					System.out.print("請輸入new加班單Word存放路徑(空白並且按下Enter取消更改): ");
					inputString = consoleInput.readLine();
					PropsHandler.setter("wordsPath", inputString);
					break;
				case 3:
					System.out.print("請輸入new產出Excel存放路徑(空白並且按下Enter取消更改): ");
					inputString = consoleInput.readLine();
					PropsHandler.setter("writePath", inputString);
					break;
				case 4:
					System.out.print("請輸入new無法處理Word存放路徑(空白並且按下Enter取消更改): ");
					inputString = consoleInput.readLine();
					PropsHandler.setter("errorWordsPath", inputString);
					break;
				case 5:
					System.out.print("請輸入new通訊錄Excel路徑(空白並且按下Enter取消更改): ");
					inputString = consoleInput.readLine();
					PropsHandler.setter("contactPath", inputString);
					break;
				case 6:
					System.out.print("請輸入new Pink Index(你調皮偷亂改後果很嚴重喔): ");
					int inputInt = Integer.parseInt(consoleInput.readLine());
					PropsHandler.setter("pinkValue", Integer.toString(inputInt));
					break;
				case 7:
					helloMsg();
					break;
				case 8:
					break changeloop;
				default:
					System.out.println("問號,你是有看到你選的這個選項??");
					
        	}
        }

		System.out.println("\n 產出Excel....");
    }
    
    
    /**
     * 歡迎訊息
     * @throws IOException
     */
    public void helloMsg() throws IOException {
        System.out.println( "Hi! " + PropsHandler.getter("userName") );
        System.out.println( "信箱 : " + PropsHandler.getter("emailAccount") );
        System.out.println( "密碼 : " + PropsHandler.getter("emailPassWord") );
        System.out.println( "加班單word存放路徑  : " + PropsHandler.getter("wordsPath") );
        System.out.println( "產出Excel會在 : " + PropsHandler.getter("writePath")  + "底下");
        System.out.println( "無法處理的特殊(機車)Word會存到 : " + PropsHandler.getter("errorWordsPath")  + "底下");
        System.out.println( "通訊錄Excel(如果加班單沒有截圖,尋找那個人的Email用) : " + PropsHandler.getter("contactPath")  + "底下");
        System.out.println( PropsHandler.getter("pinkValue") + "\n" );
    }

    /**
     * 取得加班單資料,並且計算完加班單裡面沒有的欄位資訊
     * @return
     * @throws IOException
     * @throws WordFileException 
     * @throws StopProgramException 
     */
    public List<DataPojo> returnAllWordData() throws IOException, WordFileException, StopProgramException{
    	
    	List<DataPojo> stackList = new ArrayList<>() ;

        // word 存放路徑
    	String wordsPath =  PropsHandler.getter("wordsPath") ;
    	
    	// 取得所有這路徑底下的 以.docx結尾之檔案
    	WordsFileFilter fileFilter = new WordsFileFilter() ;
    	File dir = new File(wordsPath) ;
    	File[] files = dir.listFiles(fileFilter);
    	if( files.length == 0 )
    		throw new StopProgramException("現在" + wordsPath + "底下沒有Word文件") ;
    	else {
    		for( File aFile: files ) {
    
    			// 取得 word 裡面資料 並處理一些計算的部分
    			DataPojo readyDataPojo = WordReader.readWord2007Docx(wordsPath+aFile.getName()) ;
    			// 如果不為null 初始化其中一些欄位（稍後靠PinkPojo資料對應輸入)
    			if( readyDataPojo != null ) stackList.add(readyDataPojo) ;
    			
    		}	
    	}
    	
    	if( CollectionUtils.isNotEmpty( stackList ) ) {
        	// 將 stackList 排序, 部門 > 名稱 > 使用方式 > 日期 
            Collections.sort(stackList,
            	      new Comparator<DataPojo>() {
            	          public int compare(DataPojo o1, DataPojo o2) {
            	        	// 部門 相同
            	        	if( o1.getApartment().compareTo(o2.getApartment()) == 0 ) {
            	        		// 名稱相同, 按照使用方式
                   	          	if( o1.getName().compareTo(o2.getName()) == 0 ) {
                   	          		// 使用方式相同, 按日期決定
                   	          		if(o1.getRestOrMoney().compareTo(o2.getRestOrMoney()) == 0 ) {
                   	          			return o1.getStartDay().compareTo( o2.getStartDay()) ;
                   	          		}
                   	          		
                	          		return o1.getRestOrMoney().compareTo(o2.getRestOrMoney()) ;
                	          	}
                	            
                	            return o1.getName().compareTo(o2.getName());
            	        	}
            	        	// 部門不相同
            	        	return o1.getApartment().compareTo(o2.getApartment()) ;
            	          }
            	      });
    	}else {
    		throw new StopProgramException( wordsPath + "底下沒有成功取得至少一筆Word資料!") ;
    	}
        
    	return stackList ;
    }
    
    /**
     * 寫Excel檔案到指定路徑
     * @param workbook
     * @param writePath
     * @throws IOException 
     */
    public void writeExcel(List<DataPojo> dataPojos, List<PinkPojo> pinkPojos) throws IOException {
    	
        // word 存放路徑
    	String writePath =  PropsHandler.getter("writePath") ;
    	
    	try {
    		Workbook workbook = ExcelWriter.exportData(dataPojos, pinkPojos) ; // POI會幫我們處理所有格式上所需
            FileOutputStream out=new FileOutputStream(writePath); 
    		workbook.write(out);
    		System.out.println("建立Excel成功\n");
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
    
    /**
     * 比對震旦雲原檔資料和加班單資料, 如果有對應的, 就計算時間並且加入到excel表格1後面
     * 沒有的則放入 表格2
     * 有加班單 但是上下班沒有打卡 --> 實際上下班時間 無
     * 呼叫writeExcel(dataPojos);
     * @param pinkPojos
     * @param dataPojos
     * @throws ParseException
     * @throws IOException
     */
    public void finalProcess(List<PinkPojo> pinkPojos , List<DataPojo> dataPojos) throws ParseException, IOException {
    	
    	List<PinkPojo> removePinks = new ArrayList<>() ; 
    	
    	// for 每一筆dataPojo, 並且尋找(By 名字 日期 )有沒有對應的pinkPojo
    	for( DataPojo data: dataPojos ) {
        	for( PinkPojo pinkPojo : pinkPojos ) {
        		// 有的話計算相關資料並且更新dataPojo對應欄位, 並且把該 pinkPojo 從 pinkPojos移除
        		if( pinkNameHandle( pinkPojo.getEmployee() ).equals(data.getName())  
        			&& pinkDateHandle( pinkPojo.getDate() ).equals(data.getStartDay()) ) {

        			data.setActualStartTime( pinkPojo.getOnTime() ); // 實際上班時間
        			data.setActualEndTime( pinkPojo.getOffTime() ); // 實際下班時間
        			
        			// 對照申請時數 ( 實際上-實際下 ) - (申請上-申請下)  , 存分鐘
        			if( pinkPojo.getOnTime() != "" && pinkPojo.getOffTime() != "" ) {
            			data.setDifferTotalTime( calDifferTime(pinkPojo.getStartHour(), pinkPojo.getStartMin(), pinkPojo.getEndHour(), pinkPojo.getEndMin())
            					- calDifferTime(data.getStartHour(), data.getStartMin(), data.getEndHour(), data.getEndMin()) );
        			}else
        				data.setDifferTotalTime( 0 );

        			// 是否需調整( 是否為星期天)
        			data.setSunday(isSunday(data.getStartDay())) ;
        			
        			// 缺卡內容
        			data.setMissContent( pinkPojo.getMissContent() );
        			
        			removePinks.add(pinkPojo) ; 
        		}
        	}
        	
    	}
    	
    	for( PinkPojo removedPinkPojo : removePinks)
    		pinkPojos.remove(removedPinkPojo) ; 
    		
    	// 保險起見 清空list裡面null 不管有沒有
    	while (dataPojos.remove(null));
    	if( CollectionUtils.isNotEmpty(pinkPojos)) while (pinkPojos.remove(null));
    	
    	writeExcel(dataPojos, pinkPojos);
    }
   
    /**
     * 寄信給忘記截圖的同仁
     * @param dataPojos
     * @throws IOException 
     */
    private void sendMail( List<DataPojo> dataPojos) throws IOException {

		BufferedReader consoleInput=new BufferedReader(new InputStreamReader(System.in));
		System.out.println("要寄信給加班單沒有截圖的人嗎？(Y/N)");

		char yesNo ;
		yesNo = consoleInput.readLine().charAt(0);

		if( yesNo == 'Y' ){
			System.out.println("\n寄信中....");
			for( DataPojo d : dataPojos ) {
				if( !d.isHasPhoto() ) {
					// send mail
					OutlookSender.sendMail(d.getName(), d.getStartDay());
				}
			}

			System.out.println("信件全數送完");
		}


    }
    
    
    
    // 處理震旦雲資料
    // 因為pink 裡面名稱會是部門＋姓名,  所以要透過這個function取出
    public String pinkNameHandle(String pinkName) {
    	
    	String ret_Str = pinkName.substring(pinkName.lastIndexOf('-') + 1, pinkName.length() ) ;
    	
    	return cleanString(ret_Str) ;
    }
    
    public String pinkDateHandle(String pinkDate) {
    	
    	// 轉成 yyyy/mm/dd 形式
    	int first = pinkDate.indexOf('-');
    	int second = pinkDate.indexOf("-", first + 1);
    	
    	String year= pinkDate.substring(0, first ) ;
    	String month= pinkDate.substring(first + 1, second ) ;
    	String date= pinkDate.substring(second + 1, pinkDate.indexOf('(') ) ;

    	return year + "/" + month + "/" + date ;
    }
    
    /**
     * 看日期是否為星期天
     * @param bDate
     * @return
     * @throws ParseException
     */
    public boolean isSunday(String bDate) throws ParseException {
    	
        DateFormat format1 = new SimpleDateFormat("yyyy/MM/dd");
        Date bdate = format1.parse(bDate);
        Calendar cal = Calendar.getInstance();
        cal.setTime(bdate);
        if( cal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
            return true ; 
        } 
        
        return false ;
    }

    /**
     * 計算兩個時間的差,回傳分鐘數
     * @param startHour
     * @param startMin
     * @param endHour
     * @param endMin
     * @return
     */
    public int calDifferTime(int startHour, int startMin, int endHour, int endMin) {
    	
    	Time start = new Time(startHour, startMin) ;
    	Time end = new Time(endHour, endMin) ;
    	
    	Time differTime = Time.diffTime(start, end);
    	
    	return differTime.getHours() * 60 + differTime.getMinutes() ;
    	
    }
    
    private static String cleanString(String unCleanString) { 
    	return unCleanString.trim().replaceAll("\r\n|\r|\n|\t|\f|\b", "").replaceAll("\\s+","")
    			.replaceAll("[　*| *| *|//s*]*", "").replaceAll("_", "") ;
    }
}
