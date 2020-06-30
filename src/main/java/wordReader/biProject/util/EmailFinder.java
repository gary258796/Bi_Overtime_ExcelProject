package wordReader.biProject.util;

import java.io.File;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class EmailFinder {
	public static String returnEmail( String name ) throws EncryptedDocumentException, IOException {
		  
		 String ret_EmailString = "" ;   
		
		
		 Workbook wb = WorkbookFactory.create(new File( PropsHandler.getter("contactPath") ));
		
		 Sheet sheet = wb.getSheetAt(0) ; 
		 
		 int rowCountLast = sheet.getLastRowNum() + 1 ; // rowCount 的最大值
		 int rowCountFirst = 6 ; // rowCount 開始值, “編號”下面
		  
		 
		 for( int i = rowCountFirst ; i <= rowCountLast ; i++ ) {
			  String dString = "D" + Integer.toString(i) ;
			  CellReference cr = new CellReference(dString);
			  Row rows = sheet.getRow(cr.getRow());
			  Cell cells = rows.getCell(cr.getCol());
			  
			  if(name.equals(cells.getStringCellValue())) {
				  String jString = "J" + Integer.toString(i) ;
				  cr = new CellReference(jString);
				  rows = sheet.getRow(cr.getRow());
				  cells = rows.getCell(cr.getCol());
				  ret_EmailString = cells.getStringCellValue() ; 
				  break ; 
			  }
		 }

		 return ret_EmailString ;
	}
	
}
