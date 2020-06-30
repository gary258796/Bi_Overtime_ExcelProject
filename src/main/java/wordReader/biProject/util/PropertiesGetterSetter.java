package wordReader.biProject.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.commons.configuration.ConfigurationException;
import org.apache.commons.configuration.PropertiesConfiguration;

public class PropertiesGetterSetter {


	static String fileName = "application.properties"  ;
	
	/**
	 * Int 記得要轉String
	 * @param propertyName
	 * @param value
	 */
	public void changeProperties( String propertyName, String value) {
		
		
//		try {
//			FileInputStream in = new FileInputStream(getClass().getClassLoader().getResource(fileName).getFile());
//			Properties props = new Properties();
//			props.load(in);
//			in.close();
//			
//			props.setProperty("abc", value);
//			
//			FileOutputStream out = new FileOutputStream(getClass().getClassLoader().getResource(fileName).getFile());
//			props.store(out, "This description goes to the header of a file");
//			out.close();
//		} catch (FileNotFoundException e1) {
//			// TODO Auto-generated catch block
//			e1.printStackTrace();
//		} catch (IOException e) {
//			// TODO: handle exception
//		}



		
        File propertiesFile = new File(getClass().getClassLoader().getResource(fileName).getFile());   
		try {
			PropertiesConfiguration config = new PropertiesConfiguration(fileName);
	        config.setProperty(propertyName, value);
	        config.save();
		} catch (ConfigurationException e) {
			System.out.println("error : " + e);
		}           

		
	}
	

	
	public void getPropValues() throws IOException {
		
		InputStream inputStream = null;
			
		try {
			
			Properties prop = new Properties();

			inputStream = getClass().getClassLoader().getResourceAsStream(fileName);

			prop.load(inputStream);

			// get the property value and print it out
			System.out.println("pinkValue=" + PropsHandler.getter("pinkValue") );


		} catch (FileNotFoundException ex) {
			System.err.println("Property file '" + fileName + "' not found in the classpath");
			ex.printStackTrace();
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (Exception ex) {
					ex.printStackTrace();
				}
			}
		}
	}
	
}
