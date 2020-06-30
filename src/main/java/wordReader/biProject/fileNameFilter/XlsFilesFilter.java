package wordReader.biProject.fileNameFilter;

import java.io.File;
import java.io.FilenameFilter;

public class XlsFilesFilter implements FilenameFilter {

	@Override
	public boolean accept(File dir, String name) {
	
		if( name.endsWith(".xls") && !(name.startsWith("~") || name.startsWith(".") ) )
			return true ; 
		
		return false;
	}

}
