package wordReader.biProject.excelFormat;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/* Cell 的不同Style */
public class CusCellStyle {

	   /**
	    * 對齊中央, 邊匡設定為thin 黑色, 背景顏色 紫色, 粗體字
	    * @param workbook 工作簿對象
	    * @return 單元格樣式對象
	    */
	    @SuppressWarnings("deprecation")
		public static CellStyle buildHeadCellStyle(Workbook workbook) {
	     XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
	   	 //對齊方式設置
	   	 style.setAlignment(HorizontalAlignment.CENTER);
	   	 style.setVerticalAlignment(VerticalAlignment.CENTER);
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
	   	 style.setFillForegroundColor(new XSSFColor(new java.awt.Color(231,231,255)));          
	   	 style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	   	 //粗體字設置
	   	 Font font = workbook.createFont();
	   	 font.setBold(true);
	   	 style.setFont(font);
	   	 return style;
	    }   
	    /**
	     * 對齊中央, only上下邊匡設定為thin 黑色, 背景顏色
	     * @param workbook
	     * @return
	     */
	    @SuppressWarnings("deprecation")
		public static CellStyle secondHeadStyle(Workbook workbook) {
	     XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
	   	 //對齊方式設置
	   	 style.setAlignment(HorizontalAlignment.CENTER);
	   	 style.setVerticalAlignment(VerticalAlignment.CENTER);;
	   	 //邊框顏色和寬度設置
	   	 
	   	 style.setBorderBottom(BorderStyle.THIN);
	   	 style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下邊框
	   	 style.setBorderTop(BorderStyle.THIN);
	   	 style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上邊框
	   	 
	   	 //設置背景顏色
	   	 style.setFillForegroundColor(new XSSFColor(new java.awt.Color(252,228,214)));          
	   	 style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	   	 style.setWrapText(true);
	   	 
	   	 Font font = workbook.createFont();
	   	 font.setFontName("Calibri");
	   	 style.setFont(font);

	   	 return style;
	    }
	    
	    /**
	     * 對齊中央, 邊匡設定為thin 黑色, 粗體字
	     * @param workbook
	     * @return
	     */
	    public static CellStyle thirdHeadStyle(Workbook workbook) {
	     XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
	   	 //對齊方式設置
	   	 style.setAlignment(HorizontalAlignment.CENTER);
	   	 style.setVerticalAlignment(VerticalAlignment.CENTER);;
	   	 //邊框顏色和寬度設置
	   	 
	   	 style.setBorderBottom(BorderStyle.THIN);
	   	 style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下邊框
	   	 style.setBorderLeft(BorderStyle.THIN);
	   	 style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左邊框
	   	 style.setBorderRight(BorderStyle.THIN);
	   	 style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右邊框
	   	 style.setBorderTop(BorderStyle.THIN);
	   	 style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上邊框
	   	 
	   	 //粗體字設置
	   	 Font font = workbook.createFont();
	   	 font.setBold(true);
	   	 style.setFont(font);


	   	 return style;
	    }
	    
	    /**
	     * 對齊中央, 邊匡設定為thin 黑色, 粗體字
	     * @param workbook
	     * @return
	     */
	    public static CellStyle fourthHeadStyle(Workbook workbook) {
	     XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
	   	 //對齊方式設置
	   	 style.setAlignment(HorizontalAlignment.LEFT);
	   	 style.setVerticalAlignment(VerticalAlignment.CENTER);;
	   	 //邊框顏色和寬度設置
	   	 
	   	 style.setBorderBottom(BorderStyle.THIN);
	   	 style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下邊框
	   	 style.setBorderLeft(BorderStyle.THIN);
	   	 style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左邊框
	   	 style.setBorderRight(BorderStyle.THIN);
	   	 style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右邊框
	   	 style.setBorderTop(BorderStyle.THIN);
	   	 style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上邊框
	   	 
	   	 //粗體字設置
	   	 Font font = workbook.createFont();
	   	 font.setBold(true);
	   	 style.setFont(font);
	   	 style.setWrapText(true);


	   	 return style;
	    }
	    
	    /**
	     * 紅色字體  垂直水平居中 邊線設定
	     * @param workbook
	     * @return
	     */
	    public static CellStyle redFontCellStyle(Workbook workbook) {
	    	CellStyle style = workbook.createCellStyle();
	        Font font = workbook.createFont();
	        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
	        style.setFont(font);
	        style.setBorderLeft(BorderStyle.THIN); 
	        style.setBorderRight(BorderStyle.THIN);
	        style.setBorderTop(BorderStyle.THIN);
	        style.setBorderBottom(BorderStyle.THIN);
	        style.setAlignment(HorizontalAlignment.CENTER);
	        style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
	        return style ; 
	    }
	 
}
