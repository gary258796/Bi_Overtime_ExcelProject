package wordReader.biProject.action.inMainAction.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import wordReader.biProject.cusError.ExcelException;
import wordReader.biProject.model.PinkPojo;
import wordReader.biProject.util.PropsHandler;

@Getter
@Setter
public class HandlePink {

	private int bodyStart; // 日期的row
	private int dateColumn; // 日期的column
	private int employeeColumn; // 員工的column
	private int onColumn; // 上班的column
	private int offColumn; // 下班的column
	private Workbook workBook;

	public HandlePink() {
		this.bodyStart = 0;
		this.dateColumn = 0;
		this.employeeColumn = 0;
		this.onColumn = 0;
		this.offColumn = 0;
	}

	public List<PinkPojo> handlePinkExcel() throws IOException, ExcelException {

		// 取得Excel存放路徑 (和word 相同)
		String excelPath = PropsHandler.getter("wordsPath");

		// 取得這路徑底下的 以.xls結尾之檔案
		File xlsFile = ExcelHandler.getExcelFileUnderPath(excelPath);

		// 讀入該檔案
		workBook = WorkbookFactory.create(new File(excelPath + xlsFile.getName()));

		// 取得日期、員編-姓名、上班、下班的column
		findFieldsColumn();

		// 取得pink row的資料
		List<PinkPojo> pinkPoJos = findPinkRow();

		if(CollectionUtils.isNotEmpty(pinkPoJos)) {
			// 排序
			pinkPoJos.sort(Comparator.comparing(PinkPojo::getEmployee)
			 						 .thenComparing(PinkPojo::getDate));
		}

		// 回傳資料
		return pinkPoJos;
	}

	/**
	 * 尋找 日期、員工、上班、下班的欄位index 並且存到global variables
	 */
	private void findFieldsColumn() {
		Sheet sheet = workBook.getSheetAt(0);
		int rowCountLast = getLastRowWithData(sheet); // 取得WorkBook的Row數量

		int countOfFind = 0;
		for (int i = 0; i < rowCountLast; i++) { // loop from first row to last one
			Row curRow = sheet.getRow(i);  // current row
			if (curRow != null) {
				int noOfColumns = curRow.getLastCellNum(); // get cell num of row
				for (int j = 0; j < noOfColumns; j++) {    // loop through cells in a row
					Cell curCell = curRow.getCell(j);      // get current cell
					if (curCell != null) {
						// 找到日期
						switch (getCellValue(curCell)) {
							case "日期":
								countOfFind++;
								setDateColumn(j);
								// 取得body內容第一行位置, 因為日期合併上下兩格, 所以+2
								setBodyStart(i + 2);
								break;
							case "員編-姓名":
								countOfFind++;
								setEmployeeColumn(j);
								break;
							case "上班":
								countOfFind++;
								setOnColumn(j);
								break;
							case "下班":
								countOfFind++;
								setOffColumn(j);
								break;
						}

						if (countOfFind >= 4)
							return;
					}
				}
			}
		}
	}

	/**
	 * 取得所有粉紅色列的資訊存到pinkPoJos
	 * @throws IOException :
	 */
	public List<PinkPojo> findPinkRow() {
		List<PinkPojo> pinkPoJos = new ArrayList<>();
		Sheet sheet = workBook.getSheetAt(0);
		int rowCountLast = getLastRowWithData(sheet); // rowCount 的最大值
		for (int i = getBodyStart(); i < rowCountLast; i++) {
			Row curRow = sheet.getRow(i);   // get current row
			if (curRow != null) {
				Cell onCell = curRow.getCell(getOnColumn());
				Cell offCell = curRow.getCell(getOffColumn());
				Cell employeeCell = curRow.getCell(getEmployeeColumn());
				Cell dateCell = curRow.getCell(getDateColumn());
				String dateVal = getCellValue(dateCell);
				// 判斷 1.是六或者日
				if (dateCell != null && isWeekend(dateVal)) {
					// 至少上/下班任一有值
					if ( !"".equals(getCellValue(onCell)) || !"".equals(getCellValue(offCell))) {
						String missContent = "";
						if ("".equals(getCellValue(onCell)))
							missContent = "上班";
						else if ("".equals(getCellValue(offCell)))
							missContent = "下班";

						PinkPojo pinkPojo = new PinkPojo();
						pinkPojo.setDate(getCellValue(dateCell));
						pinkPojo.setEmployee(getCellValue(employeeCell));
						pinkPojo.setOnTime(getCellValue(onCell));
						pinkPojo.setOffTime(getCellValue(offCell));
						pinkPojo.setMissContent(missContent);
						pinkPojo.setStartHour(handleTime(getCellValue(onCell))[0]);
						pinkPojo.setStartMin(handleTime(getCellValue(onCell))[1]);
						pinkPojo.setEndHour(handleTime(getCellValue(offCell))[0]);
						pinkPojo.setEndMin(handleTime(getCellValue(offCell))[1]);
						// Add to list
						pinkPoJos.add(pinkPojo);
					}
				}
			}
		}

		return pinkPoJos ;
	}

	/**
	 * 字串如果包含“六” or “日”，就回傳true
	 * @param dateString
	 * @return Boolean
	 */
	private boolean isWeekend(String dateString) {
		return dateString.contains("六") || dateString.contains("日");
	}

	/**
	 * 取得Cell的值, return一律轉為String
	 *
	 * @param cell:
	 * @return :
	 */
	private String getCellValue(Cell cell) {

		FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

		if (cell != null) {
			switch (evaluator.evaluateInCell(cell).getCellType()) {
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

	/**
	 * Return hour and minute in String
	 * @param time
	 * @return Int Array
	 */
	public int[] handleTime(String time) {
	
    	int[] ret_int = new int[2];
    	
    	if(!time.equals("")) { // // 不為空
    		int index = time.indexOf(":") ;
    		
        	try {
        		ret_int[0] = Integer.parseInt(time.substring(0, index)) ;
            } catch (NumberFormatException nfe) {
            	
            	int blank = time.indexOf(" ") ;
            	int hour = Integer.parseInt(time.substring(blank+1, index)) ;
            	ret_int[0] = hour ;
            	
            }

        	ret_int[1] = Integer.parseInt(time.substring(index+1)) ;
    	}

		return ret_int;
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
		if (r != null) {
			Iterator<Cell> cellIter = r.cellIterator();
			while (cellIter.hasNext()) {
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

}
