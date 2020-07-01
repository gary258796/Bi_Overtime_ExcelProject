package wordReader.biProject;


import java.text.ParseException;
import wordReader.biProject.cusError.ExcelException;


public class Start {

	public static void main(String[] args) throws ExcelException, ParseException {
		// TODO Auto-generated method stub

		App startApp = new App() ; 
		
		try {
			startApp.main();
		} catch (Exception e) {
			System.out.println( e.getMessage());
		}
		
	}

}
