package wordReader.biProject;

import java.text.ParseException;
import wordReader.biProject.cusError.ExcelException;

public class Start {

	// App Start Entry
	public static void main(String[] args) throws ExcelException, ParseException {

		App startApp = new App() ; 
		
		try {
			startApp.main();
		} catch (Exception e) {
			System.out.println( e.getMessage());
		}

	}

}
