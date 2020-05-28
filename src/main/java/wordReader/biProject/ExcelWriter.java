package wordReader.biProject;

import java.util.ArrayList;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument; 
import org.apache.poi.hwpf.usermodel.Range; 
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.TableShape;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor; 
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import fr.opensagres.xdocreport.utils.StringEscapeUtils;

public class ExcelWriter {
 
	 private static List<String> CELL_HEADS; //列頭
	 
	// 欄位名稱
	 static{ 
		 CELL_HEADS = new ArrayList<>();
		 CELL_HEADS.add("編號");
		 CELL_HEADS.add("部門");
		 CELL_HEADS.add("員工姓名");
		 CELL_HEADS.add("申請時間");
		 CELL_HEADS.add("(起)");
		 CELL_HEADS.add("申請時間");
		 CELL_HEADS.add("(迄)");
		 CELL_HEADS.add("申請日期");
		 CELL_HEADS.add("申請時數");
		 CELL_HEADS.add("專案");
		 CELL_HEADS.add("事由");
		 CELL_HEADS.add("承認工時");
		 CELL_HEADS.add("使用方式");
		 CELL_HEADS.add("備註");
		 CELL_HEADS.add("出勤");
	 }
	 
 
	 /**
	  * 生成Excel並寫入數據信息 ,這個function等於是ExcelWriter 的 Main Class
	  * @param dataList 數據列表
	  * @return 寫入數據後的工作簿對象
	  */
	  public static Workbook exportData(List<DataPojo> dataList){
	 	 // 生成xlsx的Excel
	 	 Workbook workbook = new SXSSFWorkbook();
	  
	 	 // 如需生成xls的Excel，請使用下面的工作簿對象，注意後續輸出時文件後綴名也需更改為xls
	 	 //Workbook workbook = new HSSFWorkbook();
	  
	 	 // 生成Sheet表，寫入第一行的列頭
	 	 Sheet sheet = buildDataSheet(workbook);
	 	 //構建每行的數據內容
	 	 int rowNum = 5; // 因為表頭是建立在index 4的地方
	 	 for (Iterator<DataPojo> it = dataList.iterator(); it.hasNext(); ) {
	 		 DataPojo data = it.next();
	 		 if (data == null) {
	 			 continue;
	 		 }
	 			 //輸出行數據
	 		 Row row = sheet.createRow(rowNum++);
	 		 convertDataToRow(workbook, sheet, data, row);
	 	 }
	 	 
	 	 // 合併編號儲存格
	 	 // create a function 
	 	 // run through whole 表單  判斷 C 欄位的 員工姓名是否相同
	 	 // 相同則編號都相同 , 並且當不相同時 , 把 編號 +1 繼續跑
	 	 handleIdn(sheet) ;
	 			 
	 	return workbook;
	  }
	  
	  /**
	   * 處理前面編號合併部分, 名稱部分相同的row 編號column就會合併
	   * @param sheet
	   */
	  private static void handleIdn(Sheet sheet ) {
			 
		  int columnIndex = 0; // for 尋找A欄位裡面, “編號”是從第幾格開始
		  int rowCountLast = sheet.getLastRowNum() + 1 ; // rowCount 的最大值
		  int rowCountFirst = 0 ; // rowCount 開始值, “編號”下面
		  
		  try {
			  // loop through A column and find cell with value : 編號
			  for (int rowIndex = 0; rowIndex < rowCountLast ; rowIndex++ ){
			      Row row = sheet.getRow(rowIndex);
			      if( row != null) {
				      Cell cell = row.getCell(columnIndex);
				      if( cell != null ) {
					      if( cell.getStringCellValue() == "編號" ) {
					    	  System.out.println("編號 在第" + (rowIndex+1) + "行" ) ;
					    	  rowCountFirst = rowIndex + 1 ; 
					      }  
				      }
			      }
			  }
		  } catch (Exception e) {
			  System.out.println("Cannot get a STRING value from a NUMERIC cell 可忽略這個在尋找編號欄位的錯誤.");
		  }
			
		  // 在接到第一次error之後就會跳出for迴圈
		  // now, 從rowCountFirst到rowCountLast, 找出Column C裡面名字相同的row並把A欄位合併
		  columnIndex = 2 ; // 尋找Column C 
		  try {
			  int idn = 1 ; // 編號初始號碼
			  boolean getNameTime = true ;
			  
			  Cell baseCell = null ;
			  Cell baseIdnCell = null ;
			  String baseName = "" ;
			  int baseIndex = 0 ;
			  
			  Cell nextCell = null  ;
			  String nextName = "" ;
			  int nextIndex = 0 ;
			  
 			  for( int i = rowCountFirst ; i < rowCountLast ; i++ ) {
				
 				  if( getNameTime ) {
 					  Row row = sheet.getRow(i);
 				      if( row != null) {
 					      baseCell = row.getCell(columnIndex);
 					      if( baseCell != null ) {
 					    	  baseIdnCell = row.getCell(0);
 					    	  
 					    	  baseName = baseCell.getStringCellValue() ; // 先找到第一個名子
 					    	  baseIndex = i ; // 紀錄Index
 					    	  getNameTime = false ; // 關閉找基底名字開關
 					      }
 				      }
 				  }
 				  
 				  if( !getNameTime ) {
 					  if( (i + 1) <  rowCountLast ) {
 	 					  Row row = sheet.getRow(i + 1);
 	 				      if( row != null ) {
 	 				    	  nextCell = row.getCell(columnIndex);
 	 					      if( nextCell != null ) {
 	 					    	  nextName = nextCell.getStringCellValue() ; // 取得名字
 	 					    	  if( nextName != baseName ) { // 名字不相同
 	 					    		  
 	 					    		  nextIndex = i + 1 ;
 	 					    		  
 	 					    		  // 代表 要做合併了 從baseIndex到nextIndex
 	 							 	  CellRangeAddress cra = new CellRangeAddress(baseIndex,nextIndex,0,0) ;
 	 							 	  sheet.addMergedRegion(cra) ;
 	 							 	  // 設定編號號碼, 都設定到baseIndex那個cell去
 	 							 	  baseIdnCell.setCellValue(idn);
 	 							 	  idn = idn + 1 ; 
 	 					    		  // 合併完之後 把getNameTime重新開啟
 	 					    		  getNameTime = true ; 
 	 					    		 
 	 					    	  }
 	 					      }
 	 				      }
 					  } // end if()
 					  else {
						 CellRangeAddress cra = new CellRangeAddress(baseIndex,i,0,0) ;
						 sheet.addMergedRegion(cra) ;
						 baseIdnCell.setCellValue(idn);
 					  }
 				  }

			  }
			  
		  } catch (Exception e) {
			  System.out.println(e);
		  }
		  
	  }
 
	  /**
	   * 生成sheet表，並寫入第一行數據（列頭）and 相關說明 更新日期等
	   * @param workbook 工作簿對象
	   * @return 已經寫入列頭的Sheet
	   */
	   private static Sheet buildDataSheet(Workbook workbook) {
	  	 
	  	 Sheet sheet = workbook.createSheet(); // createSheet裡面放參數 可指定工作表名稱
	  	 int offset = 0 ; 
	  	 // 設置 Column 寬度
	  	 for (int i=0; i<CELL_HEADS.size(); i++) {
	  		 if( CELL_HEADS.get(i) == "專案") {
	  			 sheet.setColumnWidth(i, 3500);
	  			 offset = offset + 1 ; 
	  			 sheet.setColumnWidth(i + offset, 3500);
	  		 } else if( CELL_HEADS.get(i) == "事由" ) {
	  			 sheet.setColumnWidth(i + offset, 4000);
	  			 offset = offset + 1 ;
	  			 sheet.setColumnWidth(i + offset, 4000);
	  			 offset = offset + 1 ;
	  			 sheet.setColumnWidth(i + offset, 10000);
	  		 }
	  		 else if( CELL_HEADS.get(i) == "(起)" ||  CELL_HEADS.get(i) == "(迄)" ) {
	  			 sheet.setColumnWidth(i, 2500);
	  		 } else if ( CELL_HEADS.get(i) == "承認工時" ||  CELL_HEADS.get(i) == "使用方式" || CELL_HEADS.get(i) == "備註") {
	  			 sheet.setColumnWidth(i + offset, 5000);
	  		 }
	  		 else 
	  			 sheet.setColumnWidth(i, 4000);
	  	 }
	  	 // 設置默認行高
	  	 sheet.setDefaultRowHeight((short) 400);

	  	 // 標題 : 03月份員工加班明細表
	  	 // 取得範圍內的儲存格 (start row , end row , start column , end column )
	  	 CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 1, 13); 
	  	 // 合併儲存格
	  	 sheet.addMergedRegion(cellRangeAddress) ;
	  	 // 創建標題row 
	  	 Row title = sheet.createRow(0);
	  	 title.setHeight((short) 800); // title row height : 800 
	  	 Cell titleCell = title.createCell(1);
	  	 titleCell.setCellValue("03月份員工加班明細表"); // 標題
	  	 Font font = workbook.createFont() ; 
	  	 font.setFontHeightInPoints((short) 30); // 字體大小設定為30
	  	 font.setBold(true); // 粗體
	  	 CellStyle titleStyle = workbook.createCellStyle() ;
	  	 titleStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
	  	 titleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	  	 titleStyle.setFont(font); // 塞入字體風格
	  	 titleCell.setCellStyle(titleStyle); // 塞入style到cell 
	  	 
	  	 // 標題右邊說明
	  	 cellRangeAddress = new CellRangeAddress(0, 2, 14, 16); 
	  	 sheet.addMergedRegion(cellRangeAddress) ;
	  	 Cell detailCell = title.createCell(14) ;
	  	 detailCell.setCellValue("109年確認，承認工時需扣除12-13、18-19吃飯時間(半小時不作扣除)，配合以下勞動基準法 - 勞動條件及就業平等目：\n" + 
	  	 		"第 35 條-勞工繼續工作四小時，至少應有三十分鐘之休息。但實行輪班制或其工作有連續性或緊急性者，雇主得在工作時間內，另行調配其休息時間。");
	  	 Font fontSizeTenFont = workbook.createFont() ; 
	  	 fontSizeTenFont.setFontHeightInPoints((short) 10); // 字體大小設定為30
	  	 CellStyle detCellStyle = workbook.createCellStyle() ;
	  	 detCellStyle.setWrapText(true); 
	  	 detCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
	  	 detCellStyle.setFont(fontSizeTenFont);
	  	 detailCell.setCellStyle(detCellStyle);
	  	 
	  	 // 標題下面: 更新日期
	  	 cellRangeAddress = new CellRangeAddress(2,2,1,8) ;
	  	 sheet.addMergedRegion(cellRangeAddress) ;
	  	 Row updateTimeRow = sheet.createRow(2);
	  	 updateTimeRow.setHeight((short) 500);
	  	 Cell updateTimeCellString = updateTimeRow.createCell(1) ;
	  	 updateTimeCellString.setCellValue("更新日期:"); // 內容
	  	 CellStyle updateTimeCellStringStyle = workbook.createCellStyle() ;
	  	 updateTimeCellStringStyle.setAlignment(HorizontalAlignment.RIGHT); // 靠右
	  	 updateTimeCellStringStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	  	 updateTimeCellStringStyle.setFont(fontSizeTenFont); 
	  	 updateTimeCellString.setCellStyle(updateTimeCellStringStyle);
	  	 
	  	 // 更新日期
	  	 cellRangeAddress = new CellRangeAddress(2,2,9,10) ;
	  	 sheet.addMergedRegion(cellRangeAddress) ;
	  	 Cell updateTimeCell = updateTimeRow.createCell(9) ;
	  	 DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd") ;
	  	 LocalDateTime now = LocalDateTime.now() ;
	  	 updateTimeCell.setCellValue(dtf.format(now)); // 時間
	  	 CellStyle updateTimeStyle = workbook.createCellStyle() ;
	  	 updateTimeStyle.setAlignment(HorizontalAlignment.LEFT); // 靠左
	  	 updateTimeStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	  	 updateTimeStyle.setFont(fontSizeTenFont);
	  	 updateTimeCell.setCellStyle(updateTimeStyle);
	  	 

	  	 CellRangeAddress cellRangeAddress1 = new CellRangeAddress(4,4,9,10) ;
	  	 sheet.addMergedRegion(cellRangeAddress1) ;

	  	 CellRangeAddress cellRangeAddress2 = new CellRangeAddress(4,4,11,13) ;
	  	 sheet.addMergedRegion(cellRangeAddress2) ;
	  	
	  	 
	  	 // 構建頭單元格樣式
	  	 CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
	  	 // 寫入第一行各列的數據
	  	 Row head = sheet.createRow(4);
	  	 for (int i = 0; i < CELL_HEADS.size(); i++) {
	  		 if( i == 10 ) {
	  			 Cell cell = head.createCell(i+1); // 11
	  			 cell.setCellValue(CELL_HEADS.get(i));
	  			 cell.setCellStyle(cellStyle);
	  		 } else if ( i > 10 ) {
	  			 Cell cell = head.createCell(i+3); // 11
	  			 cell.setCellValue(CELL_HEADS.get(i));
	  			 cell.setCellStyle(cellStyle);
	  		 }
	  		 else {
	  			 Cell cell = head.createCell(i);
	  			 cell.setCellValue(CELL_HEADS.get(i));
	  			 cell.setCellStyle(cellStyle);
	  		 }
	  	 }
	  	 
	  	 // 補上合併儲存格的邊筐缺陷
	  	 setCombineRegionBorder(sheet, cellRangeAddress1) ; 
	  	 setCombineRegionBorder(sheet, cellRangeAddress2) ; 
	  	 
	  	 return sheet;
	   }
	   
	   /**
	    * 設置第一行列頭的樣式
	    * @param workbook 工作簿對象
	    * @return 單元格樣式對象
	    */
	    private static CellStyle buildHeadCellStyle(Workbook workbook) {
	   	 CellStyle style = workbook.createCellStyle();
	   	 //對齊方式設置
	   	 style.setAlignment(HorizontalAlignment.CENTER);
	   	 //邊框顏色和寬度設置
	   	 style.setBorderBottom(BorderStyle.THIN);
	   	 style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下邊框
	   	 style.setBorderLeft(BorderStyle.THIN);
	   	 style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左邊框
	   	 style.setBorderRight(BorderStyle.THIN);
	   	 style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右邊框
	   	 style.setBorderTop(BorderStyle.THIN);
	   	 style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上邊框
	   	 //設置背景顏色
	   	 style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	   	 style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	   	 //粗體字設置
	   	 Font font = workbook.createFont();
	   	 font.setBold(true);
	   	 style.setFont(font);
	   	 return style;
	    }
	   
	   // 合併儲存格邊筐設定
	   private static void setCombineRegionBorder(Sheet sheet, CellRangeAddress cra) {
	  	 RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
	  	 RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
	  	 RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet);
	  	 RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
	  	 RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
	  	 RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
	  	 RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
	  	 RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
	   }
 
	   /**
	    * 將數據轉換成行
	    * @param data 源數據
	    * @param row 行對象
	    * @return
	    */
	    private static void convertDataToRow(Workbook workbook, Sheet sheet , DataPojo data, Row row){
	   	 int cellNum = 0;
	   	 int offset = 0 ; 
	   	 Cell cell;
	   	 
	   	 // 編號
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(row.getRowNum()); // error ? 
	   	 
	   	 // 部門
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getApartment() == null ? "" : data.getApartment()); 
	   	 
	   	 // 員工姓名
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getName() == null ? "" : data.getName()); 
	   	 
	   	 // 申請時間
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getStartDay() == null ? "" : data.getStartDay()); 
	   	 
	   	 // (起)
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getStartTime() == null ? "" : data.getStartTime() ); 
	   	 
	   	 // 申請時間
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getEndDay() == null ? "" : data.getEndDay()); 
	   	 
	   	 // (迄)
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getEndTime() == null ? "" : data.getEndTime()); 
	   	 
	   	 // 申請日期
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getDate() == null ? "" : data.getEndTime()); 
	   	 
	   	 // 申請時數
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getApplyHour() == null ? "" : data.getApplyHour()); 
	   	 
	   	 // 專案
	   	 CellRangeAddress cra = new CellRangeAddress(row.getRowNum(),row.getRowNum(),9,10) ;
	   	 sheet.addMergedRegion(cra) ;
	   	 cell = createCellWithAlignment(workbook, row, cellNum++);
	   	 cell.setCellValue(data.getProjectName() == null ? "" : data.getProjectName()); 
	   	 offset = offset + 1 ; 
	   	 
	   	 // 事由
	   	 cra = new CellRangeAddress(row.getRowNum(),row.getRowNum(),11,13) ;
	   	 sheet.addMergedRegion(cra) ;
	   	 cell = createCellWithLeftAlignment(workbook, row, cellNum + offset);
	   	 cell.setCellValue(data.getReason() == null ? "" : data.getReason()); 
	   	 cellNum = cellNum + 1 ; 
	   	 offset = offset + 2 ; 
	   	 
	   	 // 承認工時
	   	 cell = createCellWithAlignment(workbook, row, cellNum + offset);
	   	 cell.setCellValue(data.getAdmitTime() == null ? "" : data.getAdmitTime()); 
	   	 cellNum = cellNum + 1 ;
	   	 
	   	 // 使用方式
	   	 cell = createCellWithAlignment(workbook, row, cellNum + offset);
	   	 cell.setCellValue(data.isRestOrMoney() == true ? "補修" : "加班費"); 
	   	 cellNum = cellNum + 1 ;
	   	 
	   	 // 備註
	   	 cell = createCellWithLeftAlignment(workbook, row, cellNum + offset);
	   	 cell.setCellValue("OK"); 
	   	 cellNum = cellNum + 1 ;
	   	 
	   	 // 出勤
	   	 cell = createCellWithAlignment(workbook, row, cellNum + offset);
	   	 cell.setCellValue(""); 
	   	 cellNum = cellNum + 1 ;
	   	 
	    }
 
	    // 水平 垂直居中
	    private static Cell createCellWithAlignment(Workbook workbook, Row row, int cellNum) {
	   	 Cell ret_Cell = row.createCell(cellNum);
	   	 CellStyle retCellStyle = workbook.createCellStyle() ;
	   	 retCellStyle.setAlignment(HorizontalAlignment.CENTER);
	   	 retCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	   	 ret_Cell.setCellStyle(retCellStyle);
	   	 
	   	 return ret_Cell ; 
	    }
 
	    // 垂直居中 水平靠左
	    private static Cell createCellWithLeftAlignment(Workbook workbook, Row row, int cellNum) {
	   	 Cell ret_Cell = row.createCell(cellNum);
	   	 CellStyle retCellStyle = workbook.createCellStyle() ;
	   	 retCellStyle.setAlignment(HorizontalAlignment.LEFT);
	   	 retCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	   	 ret_Cell.setCellStyle(retCellStyle);
	   	 
	   	 return ret_Cell ; 
	    }
 
}