package wordReader.biProject.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
import org.apache.commons.configuration.ConfigurationException;
import org.apache.commons.configuration.PropertiesConfiguration;
import org.apache.commons.lang.StringUtils;
import wordReader.biProject.App;

/**
 * 負責存取properties檔案
 */
public class PropsHandler {


	public static String getter(String getItem) throws IOException {

        Properties props = new Properties();

        props.load(new FileInputStream( getPropertiesPath("propertiesPath") ));
        
        return props.getProperty(getItem) ; 
	}


	public static void setter(String field, String value) throws IOException, ConfigurationException {

		// 獲得的值不等於空白 or Null才更新property的值
		if(!StringUtils.isBlank(value)){
			PropertiesConfiguration config = new PropertiesConfiguration(getPropertiesPath("propertiesPath"));

			config.setProperty(field, value);

			config.save();
		}
	}


	/**
	 * Read property from application.properties file
	 * @return :
	 * @throws IOException:
	 */
	public static String getPropertiesPath(String target) throws IOException {

		Properties props = new Properties();

		props.load(App.class.getClassLoader().getResourceAsStream("application.properties"));

		return props.getProperty(target) ;

	}
	

	
}
