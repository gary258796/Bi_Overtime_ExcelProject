package wordReader.biProject;

import java.util.List;
import org.apache.commons.collections4.CollectionUtils;
import wordReader.biProject.action.afterMainAction.AfterMain;
import wordReader.biProject.action.afterMainAction.AfterMainImpl;
import wordReader.biProject.action.beforeMainAction.BeforeMain;
import wordReader.biProject.action.beforeMainAction.BeforeMainImpl;
import wordReader.biProject.action.inMainAction.excel.HandlePink;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.model.PinkPojo;
import wordReader.biProject.util.PropsHandler;
import wordReader.biProject.action.inMainAction.word.WordHandle;
import wordReader.biProject.action.inMainAction.word.WordHandleImpl;

public class App 
{
	/** 儲存從Excel取得的震旦雲打卡資料 */
	private List<PinkPojo> pinkPojos ;
	/** 處理Excel裡面被標記為粉紅色(假日加班)之資訊Util */
	private final HandlePink handlePink = new HandlePink() ;

	// ------------------------------------------------------------------ //
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
	}

	public void main() throws Exception {
    	// 歡迎訊息 並且會顯示 word取得路徑 以及excel產生路徑...等訊息
		beforeMain.helloMsg();
    	// 詢問是否更改相關參數
    	beforeMain.changeProperties();

		List<DataPojo> dataPojos;
		try {
    		// 取得加班單資料,並且計算完加班單裡面沒有的欄位資訊
    		// 無法成功獲取的word會放入到指定路徑底下
			dataPojos = wordHandle.returnAllWordData();

        	// 取得震旦雲原檔裡面,粉紅色row的每一筆資料(上/下班欄位至少有一個是有值的)
    		pinkPojos = handlePink.handlePinkExcel();
		} catch (Exception e) {
			throw new Exception(e.getMessage()) ;
		}  finally {
    		// 檢查pinkPojos是否是空的，空的代表說設定之粉紅顏色參數有誤
	    	if( CollectionUtils.isEmpty(pinkPojos) || pinkPojos.size() == 0 )
	    		System.out.println("PinkPojos 為null或者大小為0,可能是因為這個excel粉紅色欄位index不是59.");
		}

    	// 將準備好的資料寫到 excel , Excel Writer 負責 呈現的部分
		afterMain.finalProcess(pinkPojos, dataPojos);

    	// 將DataPojo 裡面 , hasPhoto 為false的資料 送outlook mail給該位同仁
		afterMain.sendMail(dataPojos) ;

		System.out.println("程式結束.");
		System.out.println("無法正確讀取的word文件會放到妳指定的路徑: " + PropsHandler.getter("errorWordsPath"));
    }

}
