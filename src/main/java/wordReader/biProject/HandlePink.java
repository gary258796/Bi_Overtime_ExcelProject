package wordReader.biProject;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import wordReader.biProject.cusError.ExcelException;
import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.fileNameFilter.XlsFilesFilter;
import wordReader.biProject.model.PinkPojo;
import wordReader.biProject.util.PropsHandler;

public class HandlePink {

	private int bodyStart = 0; // 日期的row
	private int dateColumn = 0; // 日期的column
	private int employeeColumn = 0; // 員工的column
	private int onColumn = 0; // 上班的column
	private int offColumn = 0; // 下班的column
	
	public List<PinkPojo> handlePinkExcel() throws IOException, ExcelException, StopProgramException {

		// 取得Excel存放路徑 (和word 相同)
		String wordsPath = PropsHandler.getter("wordsPath");

		// 取得這路徑底下的 以.xls結尾之檔案
		XlsFilesFilter fileFilter = new XlsFilesFilter();
		File dir = new File(wordsPath);
		File[] files = dir.listFiles(fileFilter);
		
		if (files.length == 0)
			throw new ExcelException("路徑底下沒有任一 .xls 檔案.");
		else if (files.length == 1) { // 只接受路徑底下一個.xls file

			Workbook wb = WorkbookFactory.create(new File(wordsPath + files[0].getName()));
			
			// 取得日期、員編-姓名、上班、下班的column
			findFieldsColumn(wb);

			// 取得pink row的資料
			List<PinkPojo> pinkPojos = findPinkRow(wb);

			if( CollectionUtils.isEmpty(pinkPojos) || pinkPojos.size() == 0 ) 
				return null ;
			else {
				// 排序
		        Collections.sort(pinkPojos,
		        	      new Comparator<PinkPojo>() {
		        	          public int compare(PinkPojo o1, PinkPojo o2) {
		        	          	
		        	          	if( o1.getEmployee().compareTo(o2.getEmployee()) == 0 ) {
		        	          		// 名稱相同  按照日期先後順序
		        	          		return o1.getDate().compareTo( o2.getDate()) ;
		        	          	}
		        	          	
		        	              return o1.getEmployee().compareTo(o2.getEmployee());
		        	          }
		        	      });
		        
		        
				// 回傳資料
				return pinkPojos;
			}
		} else {
			throw new StopProgramException("路徑底下有超過一個.xls的檔案");
		}

	}

	/**
	 * 尋找 日期、員工、上班、下班的欄位index 並且存到global variables
	 * 
	 * @param wb
	 * @throws ExcelException
	 */
	public void findFieldsColumn(Workbook wb) throws ExcelException {
		
		Sheet sheet = wb.getSheetAt(0);
		int rowCountLast = getLastRowWithData(sheet); // rowCount 的最大值

		// loop through all rows
		int countOfFind = 0;
		for (int i = 0; i < rowCountLast; i++) {
			
			Row curRow = sheet.getRow(i);
			if (curRow != null) {
				int noOfColumns = curRow.getLastCellNum();
				// loop through cells in a row
				for (int j = 0; j < noOfColumns; j++) {
					Cell curCell = curRow.getCell(j);
					if (curCell != null) {
						// 找到日期
						if (getCellValue(wb, curCell).equals("日期")) {
							countOfFind++;
							setDateColumn(j);
							// 取得body內容第一行位置, 因為日期合併兩格, 所以 + 2
							setBodyStart(i + 2);
						} else if (getCellValue(wb, curCell).equals("員編-姓名")) {
							countOfFind++;
							setEmployeeColumn(j);
						} else if (getCellValue(wb, curCell).equals("上班")) {
							countOfFind++;
							setOnColumn(j);
						} else if (getCellValue(wb, curCell).equals("下班")) {
							countOfFind++;
							setOffColumn(j);
						}

						if (countOfFind >= 4)
							return;
					}
				}
			}
		}
	}

	/**
	 * 取得所有粉紅色列的資訊存到pinkPojos
	 * @param workbook
	 * @throws ExcelException
	 * @throws IOException 
	 */
	public List<PinkPojo> findPinkRow(Workbook workbook) throws ExcelException, IOException {
		
		List<PinkPojo> pinkPojos = new ArrayList<>();
		Sheet sheet = workbook.getSheetAt(0);
		int rowCountLast = getLastRowWithData(sheet); // rowCount 的最大值

		for (int i = getBodyStart(); i < rowCountLast; i++) {
			
			// 判斷 1.粉紅色 && 2.至少上/下班任一有值
			Short pinkValue = getPinkValue() ;

			Row curRow = sheet.getRow(i);
			if (curRow != null) {
				
				Cell firstCell = curRow.getCell(0);
				if (firstCell != null) {
					
					CellStyle curCellStyle = firstCell.getCellStyle();
					if (curCellStyle != null) {
					
						// // 1. 是粉紅色
						if (Short.compare(curCellStyle.getFillForegroundColor(), pinkValue) == 0) {
							// 2. 上/ 下班有值
							Cell onCell = curRow.getCell(getOnColumn());
							Cell offCell = curRow.getCell(getOffColumn());
							if (getCellValue(workbook, onCell) != "" || getCellValue(workbook, offCell) != "") {
								
								String missContent = "";
								if (getCellValue(workbook, onCell) == "")
									missContent = "上班";
								else if (getCellValue(workbook, offCell) == "")
									missContent = "下班";

								Cell dateCell = curRow.getCell(getDateColumn());
								Cell employeeCell = curRow.getCell(getEmployeeColumn());
				
								// create Entity 
								PinkPojo pinkPojo = new PinkPojo(getCellValue(workbook, dateCell),
										getCellValue(workbook, employeeCell), getCellValue(workbook, onCell),
										getCellValue(workbook, offCell), missContent,
										handleTime(getCellValue(workbook, onCell))[0], handleTime(getCellValue(workbook, onCell))[1],
										handleTime(getCellValue(workbook, offCell))[0], handleTime(getCellValue(workbook, offCell))[1]);

								// Add to list 
								pinkPojos.add(pinkPojo);
							}

						}
					}
				}
			}
		}
		
		if( CollectionUtils.isEmpty(pinkPojos) || pinkPojos.size() == 0 )
			pinkPojos.clear() ;
		
		return pinkPojos ;
	}
	
	public int[] handleTime(String time) {
	
    	int[] ret_int = new int[2];
    	
    	if( time == "" ) { // 為空
        	ret_int[0] = 0 ;
        	ret_int[1] = 0 ;  
    	}else { // 不為空
    		int index = time.indexOf(":") ;
    		
        	try {
        		ret_int[0] = Integer.parseInt(time.substring(0, index)) ;
            } catch (NumberFormatException nfe) {
            	
            	int blank = time.indexOf(" ") ;
            	int hour = Integer.parseInt(time.substring(blank+1, index)) ;
            	ret_int[0] = hour ;
            	
            }
    		
    		
        	ret_int[1] = Integer.parseInt(time.substring(index+1, time.length())) ;  
    	}

		return ret_int;
	}

	/**
	 * 取得Cell的值, return一律轉為String
	 * 
	 * @param wb
	 * @param cell
	 * @return
	 * @throws ExcelException
	 */
	public String getCellValue(Workbook wb, Cell cell) throws ExcelException {

		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

		if (cell != null) {
			switch (evaluator.evaluateInCell(cell).getCellType()) {
			case BLANK:
				return "";
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			case BOOLEAN:
				System.out.println("目前沒有處理Excel裡面 BOOLEAN的欄位.");
				break;
			case ERROR:
				System.out.println("目前沒有處理Excel裡面 ERROR的欄位.");
				break;
			case FORMULA:
				System.out.println("目前沒有處理Excel裡面 FORMULA的欄位.");
				break;
			case _NONE:
				System.out.println("目前沒有處理Excel裡面 _NONE的欄位.");
				break;
			default:
				return "";
			}
		}

		return "";
	}

	
	private short getPinkValue() throws IOException {
		
		// 取得Excel存放路徑 (和word 相同)
		String pinkString = PropsHandler.getter("pinkValue");
		
		int value = Integer.valueOf(pinkString) ; 
		Short pinkNum = (short) value;
		Short pinkValue = new Short(pinkNum);
		
		return pinkValue ;
	}
	
	
	
	
	// 取得最後一行index
	public int getLastRowWithData(Sheet sheet) {
		int rowCount = 0;
		Iterator<Row> iter = sheet.rowIterator();

		while (iter.hasNext()) {
			Row r = iter.next();
			if (!isRowBlank(r)) {
				rowCount = r.getRowNum();
			}
		}

		return rowCount;
	}

	// 判斷row or cell 空
	public boolean isRowBlank(Row r) {
		boolean ret = true;

		/*
		 * If a row is null, it must be blank.
		 */
		if (r != null) {
			Iterator<Cell> cellIter = r.cellIterator();
			/*
			 * Iterate through all cells in a row.
			 */
			while (cellIter.hasNext()) {
				/*
				 * If one of the cells in given row contains data, the row is considered not
				 * blank.
				 */
				if (!isCellBlank(cellIter.next())) {
					ret = false;
					break;
				}
			}
		}

		return ret;
	}

	public boolean isCellBlank(Cell c) {
		return (c == null || c.getCellType() == CellType.BLANK);
	}

	public boolean isCellEmpty(Cell c) {
		return (c == null || c.getCellType() == CellType.BLANK
				|| (c.getCellType() == CellType.STRING && c.getStringCellValue().isEmpty()));
	}

	// Getter and Setter
	public int getDateColumn() {
		return dateColumn;
	}

	public void setDateColumn(int dateColumn) {
		this.dateColumn = dateColumn;
	}

	public int getEmployeeColumn() {
		return employeeColumn;
	}

	public void setEmployeeColumn(int employeeColumn) {
		this.employeeColumn = employeeColumn;
	}

	public int getOnColumn() {
		return onColumn;
	}

	public void setOnColumn(int onColumn) {
		this.onColumn = onColumn;
	}

	public int getOffColumn() {
		return offColumn;
	}

	public void setOffColumn(int offColumn) {
		this.offColumn = offColumn;
	}

	public int getBodyStart() {
		return bodyStart;
	}

	public void setBodyStart(int bodyStart) {
		this.bodyStart = bodyStart;
	}

}
