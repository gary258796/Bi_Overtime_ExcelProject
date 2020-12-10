package wordReader.biProject.action.inMainAction.word;

import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.cusError.WordFileException;
import wordReader.biProject.model.DataPojo;

import java.io.IOException;
import java.util.List;

public interface WordHandle {

    /**
     * 取得加班單資料,並且計算完加班單裡面沒有的欄位資訊
     * @return List of data from words
     */
    List<DataPojo> returnAllWordData() throws IOException, WordFileException, StopProgramException;

}
