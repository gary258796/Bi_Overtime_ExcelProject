package wordReader.biProject.action.inMainAction.excel;

import wordReader.biProject.cusError.ExcelException;
import wordReader.biProject.fileNameFilter.XlsFilesFilter;
import java.io.File;

/**
 * 獲取指定路徑底下的Excel檔案
 */
public class ExcelHandler {

    /**
     * Return files under @path, which are .xls format and not start with '~' and '.'
     *
     * @param path
     * @return File array
     */
    public static File getExcelFileUnderPath(String path) throws ExcelException {
        // 文件過濾器(只接受 .xls結尾並且不是'~'和'.'開頭的文件)
        XlsFilesFilter fileFilter = new XlsFilesFilter();
        // 取得該路徑資料夾
        File dir = new File(path) ;
        // 用文件過濾器取出符合的文件
        File[] validFiles = dir.listFiles(fileFilter);

        File returnFile;
        if( validFiles != null && validFiles.length == 1)
            returnFile = validFiles[0];
        else {
            throw new ExcelException("路徑底下.xls檔案需要恰恰為一個.");
        }

        return returnFile;
    }

}
