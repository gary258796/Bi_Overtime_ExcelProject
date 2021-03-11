package wordReader.biProject;

import java.util.List;
import org.apache.commons.collections4.CollectionUtils;
import wordReader.biProject.action.afterMainAction.AfterMain;
import wordReader.biProject.action.afterMainAction.AfterMainImpl;
import wordReader.biProject.action.beforeMainAction.BeforeMain;
import wordReader.biProject.action.beforeMainAction.BeforeMainImpl;
import wordReader.biProject.action.inMainAction.excel.HandlePink;
import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.model.PinkPojo;
import wordReader.biProject.util.PropsHandler;
import wordReader.biProject.action.inMainAction.word.WordHandle;
import wordReader.biProject.action.inMainAction.word.WordHandleImpl;

public class App 
{
	// ------------------------------------------------------------------ //
	/** 處理Excel裡面被標記為粉紅色(假日加班)之資訊Util */
	private final HandlePink handlePink;
	/** 執行程式之前相關: 歡迎訊息、參數設定 */
	private final BeforeMain beforeMain;
	/** 負責取得word資料 */
	private final WordHandle wordHandle;
	/** 寫入Excel並且寄信之後續動作 */
	private final AfterMain afterMain;

	/** Constructor */
	public App() {
		this.beforeMain = new BeforeMainImpl();
		this.wordHandle = new WordHandleImpl();
		this.afterMain = new AfterMainImpl();
		this.handlePink = new HandlePink() ;
	}

	public void main() throws Exception {
    	// 歡迎訊息 並且會顯示 word取得路徑 以及excel產生路徑...等訊息
		beforeMain.helloMsg();
    	// 詢問是否更改相關參數
    	beforeMain.changeProperties();

		System.out.println("\n 產出Excel ing....");

		// 加班單彙整的資料
		List<DataPojo> dataPoJos;
		// 儲存從Excel取得的震旦雲打卡資料
		List<PinkPojo> pinkPoJos;
		try {
    		// 取得加班單資料,並且計算完加班單裡面沒有的欄位資訊,無法成功獲取的word會放入到指定路徑底下
			dataPoJos = wordHandle.returnAllWordData();
        	// 取得震旦雲原檔裡面,粉紅色row的每一筆資料(上/下班欄位至少有一個是有值的)
    		pinkPoJos = handlePink.handlePinkExcel();
		} catch (Exception e) {
			throw new Exception(e.getMessage()) ;
		}

		if( CollectionUtils.isEmpty(dataPoJos))
		    throw new StopProgramException("沒有成功取得至少一筆Word資料! 請確認該路徑底下有加班單資訊的word文件們.") ;
		if( CollectionUtils.isEmpty(pinkPoJos) ) // TODO: 到時候不會抓取很紅色
			System.out.println("抓取不到任何震旦雲原檔假日資料,可能是因為這個excel粉紅色欄位index不是59.");

    	// 將準備好的資料寫到 excel , Excel Writer 負責 呈現的部分
		afterMain.finalProcess(pinkPoJos, dataPoJos);
    	// 將DataPojo 裡面 , hasPhoto 為false的資料 送outlook mail給該位同仁
		afterMain.sendMail(dataPoJos) ;

		System.out.println("程式結束.");
		System.out.println("無法正確讀取的word文件會放到指定的路徑: " + PropsHandler.getter("errorWordsPath"));
    }

}
