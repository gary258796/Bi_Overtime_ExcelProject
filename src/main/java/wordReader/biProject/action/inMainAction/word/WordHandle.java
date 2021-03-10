package wordReader.biProject.action.inMainAction.word;

import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.cusError.WordFileException;
import wordReader.biProject.model.DataPojo;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * 處理加班單word相關功能
 */
public interface WordHandle {

    /**
     * 取得加班單資料,並且計算完加班單裡面沒有的欄位資訊
     * @return List of data from words
     */
    List<DataPojo> returnAllWordData() throws IOException, WordFileException, StopProgramException;

    /**
     * Return files under @path, which are docx format and not start with '~' and '.'
     * @param path
     * @return File array
     */
    List<File> getDocxFileUnderPath(String path);




}
