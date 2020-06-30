package wordReader.biProject.fileNameFilter;

import java.io.File;
import java.io.FilenameFilter;

public class WordsFileFilter implements FilenameFilter {

	@Override
	public boolean accept(File dir, String name) {

		/**
		 * 只接受 .docx結尾 並且不是～以及.開頭的
		 */
		if( name.endsWith(".docx") && !(name.startsWith("~") || name.startsWith(".")) ) {
			// .docx結尾的是我們要的工作單
			return true;
		}
		
		return false;
	}
	
	
	

}
