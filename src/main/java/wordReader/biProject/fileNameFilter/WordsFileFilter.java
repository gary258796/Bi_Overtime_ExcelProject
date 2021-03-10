package wordReader.biProject.fileNameFilter;

import java.io.File;
import java.io.FilenameFilter;

/**
 * 過濾資料夾底下的文件, 只接受 .docx結尾並且不是～以及.開頭的
 */
public class WordsFileFilter implements FilenameFilter {

	/**
	 * Accept file with docx format and not start with '~' and '.'
	 * @param dir
	 * @param name
	 * @return
	 */
	@Override
	public boolean accept(File dir, String name) {
		return name.endsWith(".docx") && !(name.startsWith("~") || name.startsWith("."));
	}

}
