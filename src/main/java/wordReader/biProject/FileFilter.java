package wordReader.biProject;

import java.io.File;
import java.io.FilenameFilter;

public class FileFilter implements FilenameFilter {

	@Override
	public boolean accept(File dir, String name) {

		if( name.endsWith(".docx") ) {
			// .docx結尾的是我們要的工作單
			return true;
		}
		
		return false;
	}
	
	
	

}
