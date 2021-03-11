package wordReader.biProject.action.inMainAction.word;

import java.io.*;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import wordReader.biProject.action.inMainAction.util.InvalidFileHandler;
import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.util.PropsHandler;

public class WordReader {

	// TODO: get property from properties file, maybe convert to write log
	static boolean debug = false ;  // 要印出訊息的時候就設定為true , 平常是關起來false狀態

    /**
     * 取得word裡面我們要的資訊, 計算額外欄位資訓, 存到並回傳 DataPojo
     * @param docxFilePath : 目標檔案的路徑
     * @return
     * @throws IOException 
     */
    @SuppressWarnings("resource")
	public static DataPojo readWord2007Docx(String docxFilePath) throws IOException{

    	DataPojo dataPojo = null ;

   	 	try {
   	       // 取得該目標檔案word的內容
			File file = new File(docxFilePath);
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());
			XWPFDocument document = new XWPFDocument(fis);
			XWPFWordExtractor extractor = new XWPFWordExtractor(document);
			// 獲取word裡面的內容
			String bodyString =  extractor.getText() ;
	 	    // debug
		    debugOutput(debug, document.getDocument().toString(), bodyString);
   	 	    // translate data to dataPojo
   	 	    dataPojo = stringToDataPojoHelper(bodyString, document.getDocument().toString()) ;
   	 	    // close file input stream
            fis.close();
		} catch(Exception ex) {
			// 設定錯誤文件放置路徑
			String destinationOfInvalidFile = PropsHandler.getter("errorWordsPath") ;
			// 問題檔案丟到設定路徑底下
			InvalidFileHandler.throwInvalidFilePathFileToDestination(docxFilePath, destinationOfInvalidFile );
		}

   	 	return dataPojo;
    }

	/**
	 * 將word資料轉換成DataPojo儲存
	 * @param bodyString
	 * @param documentString
	 * @return
	 */
    private static DataPojo stringToDataPojoHelper(String bodyString, String documentString) throws StopProgramException {
		FieldExtractHandler fieldExtractHandler = new FieldExtractHandler(bodyString);

    	DataPojo dataPojo = new DataPojo() ;
		dataPojo.setDate( fieldExtractHandler.getApplyDate() );
		dataPojo.setApartment( fieldExtractHandler.getApartment() );
		dataPojo.setName( fieldExtractHandler.getStaffName() );
		dataPojo.setStartDay(fieldExtractHandler.getActualDate());
		dataPojo.setStartTime(fieldExtractHandler.getStartTime());
		dataPojo.setEndTime(fieldExtractHandler.getEndTime());
		dataPojo.setProjectName(fieldExtractHandler.getProjectName());
		dataPojo.setReason( fieldExtractHandler.getLateReason() );
		dataPojo.setRestOrMoney(fieldExtractHandler.restOrMoney());
		dataPojo.setExtraMsg(fieldExtractHandler.getExtraMsg());
		dataPojo.setHasPhoto(fieldExtractHandler.isImageInOrNot(documentString)) ;
		dataPojo.setEndDay(dataPojo.getStartDay());
		dataPojo.setStartHour(fieldExtractHandler.getStartHour());
		dataPojo.setStartMin(fieldExtractHandler.getStartMinute());
		dataPojo.setEndHour(fieldExtractHandler.getEndHour());
		dataPojo.setEndMin(fieldExtractHandler.getEndMinute());
		dataPojo.setApplyHour(fieldExtractHandler.getApplyHour());
		dataPojo.setAdmitTime(fieldExtractHandler.getAdmitTime());
		// 底下預設
		dataPojo.setActualStartTime("") ;
		dataPojo.setActualEndTime(""); // 實際下班時間
		dataPojo.setDifferTotalTime(0);
		dataPojo.setSunday(false) ;
		dataPojo.setMissContent("");

		return dataPojo;
	}

	/**
	 * 當Debug時，印出訊息到console
	 * @param isDebug
	 * @param documentString
	 * @param bodyString
	 */
	private static void debugOutput(boolean isDebug, String documentString, String bodyString) throws StopProgramException {
		if( isDebug ) {
			FieldExtractHandler fieldExtractHandler = new FieldExtractHandler(bodyString);

			System.out.println("bodyString " + bodyString + "\n");
			System.out.println("Date : " + fieldExtractHandler.getApplyDate());
			System.out.println("Apartment : " + fieldExtractHandler.getApartment());
			System.out.println("Name : " + fieldExtractHandler.getStaffName());
			System.out.println("StartDate : " + fieldExtractHandler.getActualDate());
			System.out.println("StartTime : " + fieldExtractHandler.getStartTime());
			System.out.println("EndTime : " + fieldExtractHandler.getEndTime());
			System.out.println("ProjectName : " + fieldExtractHandler.getProjectName());
			System.out.println("Late Reason : " + fieldExtractHandler.getLateReason());
			System.out.println("Rest or Money : " + fieldExtractHandler.restOrMoney());
			System.out.println("Extra Msg : " + fieldExtractHandler.getExtraMsg());
			System.out.println("Has Photo : " + fieldExtractHandler.isImageInOrNot(documentString));
		}
	}

}
