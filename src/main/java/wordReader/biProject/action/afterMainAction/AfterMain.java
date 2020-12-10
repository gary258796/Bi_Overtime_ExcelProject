package wordReader.biProject.action.afterMainAction;

import wordReader.biProject.model.DataPojo;
import wordReader.biProject.model.PinkPojo;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;

public interface AfterMain {

    /**
     * 比對震旦雲原檔資料和加班單資料, 如果有對應的, 就計算時間並且加入到excel表格1後面
     * 沒有的則放入 表格2
     * 有加班單 但是上下班沒有打卡 --> 實際上下班時間 無
     * 呼叫writeExcel(dataPojos)
     * @param pinkPojos :
     * @param dataPojos :
     */
    void finalProcess(List<PinkPojo> pinkPojos , List<DataPojo> dataPojos) throws ParseException, IOException;


    /**
     * 寄信給忘記截圖的同仁
     * @param dataPojos :
     */
    void sendMail( List<DataPojo> dataPojos) throws IOException;
}
