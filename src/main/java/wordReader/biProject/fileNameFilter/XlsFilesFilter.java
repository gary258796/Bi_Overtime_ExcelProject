package wordReader.biProject.fileNameFilter;

import java.io.File;
import java.io.FilenameFilter;

public class XlsFilesFilter implements FilenameFilter {

	@Override
	public boolean accept(File dir, String name) {
		return name.endsWith(".xls") && !(name.startsWith("~") || name.startsWith("."));
	}

}
