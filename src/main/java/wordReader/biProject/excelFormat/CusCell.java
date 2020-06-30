package wordReader.biProject.excelFormat;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class CusCell {

	
    /**
     * 水平 垂直居中 邊線設定
     * @param workbook
     * @param row
     * @param cellNum  
     * @param yellowBack 標記黃色背景
     * @return
     */
    @SuppressWarnings("deprecation")
	public static Cell createCellWithAlignment(Workbook workbook, Row row, int cellNum, boolean yellowBack) {
   	 Cell ret_Cell = row.createCell(cellNum);
   	 XSSFCellStyle retCellStyle = (XSSFCellStyle)workbook.createCellStyle();
//   	 CellStyle retCellStyle = workbook.createCellStyle() ;
   	 retCellStyle.setBorderLeft(BorderStyle.THIN); 
   	 retCellStyle.setBorderRight(BorderStyle.THIN);
   	 retCellStyle.setBorderTop(BorderStyle.THIN);
   	 retCellStyle.setBorderBottom(BorderStyle.THIN);
   	 retCellStyle.setAlignment(HorizontalAlignment.CENTER);
   	 retCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
   	 if(yellowBack) {
   		retCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,255,204)));          
   		retCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
   	 }
   	 Font baseFont = workbook.createFont();
   	 baseFont.setFontName("微軟正黑體");
   	 retCellStyle.setFont(baseFont);
   	 ret_Cell.setCellStyle(retCellStyle);
   	 return ret_Cell ; 
    }

    /**
     * 垂直居中 水平靠左
     * @param workbook
     * @param row
     * @param cellNum
     * @param yellowBack
     * @return
     */
    @SuppressWarnings("deprecation")
	public static Cell createCellWithLeftAlignment(Workbook workbook, Row row, int cellNum, boolean yellowBack) {
	   	 Cell ret_Cell = row.createCell(cellNum);
	   	 XSSFCellStyle retCellStyle = (XSSFCellStyle)workbook.createCellStyle();
//	   	 CellStyle retCellStyle = workbook.createCellStyle() ;
	   	 retCellStyle.setBorderLeft(BorderStyle.THIN); 
	   	 retCellStyle.setBorderRight(BorderStyle.THIN);
	   	 retCellStyle.setBorderTop(BorderStyle.THIN);
	   	 retCellStyle.setBorderBottom(BorderStyle.THIN);
	   	 retCellStyle.setAlignment(HorizontalAlignment.LEFT);
	   	 retCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	   	 if(yellowBack) {
	   		retCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,255,204)));          
	   		retCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	   	 }
	   	 
	   	 Font baseFont = workbook.createFont();
	   	 baseFont.setFontName("微軟正黑體");
	   	 retCellStyle.setFont(baseFont);
	   	 ret_Cell.setCellStyle(retCellStyle);
	   	 return ret_Cell ; 
    }
    

    
}

