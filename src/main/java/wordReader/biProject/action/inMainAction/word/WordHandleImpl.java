package wordReader.biProject.action.inMainAction.word;

import org.apache.commons.collections4.CollectionUtils;
import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.fileNameFilter.WordsFileFilter;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.util.PropsHandler;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WordHandleImpl implements WordHandle{

    /**
     * 回傳加班單word獲取的所有資料
     * @return List of data of 加班
     * @throws IOException :
     * @throws StopProgramException :
     */
    public List<DataPojo> returnAllWordData() throws IOException, StopProgramException {

        List<DataPojo> stackList = new ArrayList<>() ;

        // word 存放路徑
        String wordsPath =  PropsHandler.getter("wordsPath") ;
        // 取得所有這路徑底下的以.docx結尾之檔案
        WordsFileFilter fileFilter = new WordsFileFilter() ;
        File dir = new File(wordsPath) ;
        File[] files = dir.listFiles(fileFilter);
        if( files == null || files.length == 0 )
            throw new StopProgramException("現在" + wordsPath + "底下沒有Word文件") ;
        else {
            for( File aFile: files ) {
                // 取得 word 裡面資料 並處理一些計算的部分
                DataPojo readyDataPojo = WordReader.readWord2007Docx(wordsPath+aFile.getName()) ;
                // 如果不為null 初始化其中一些欄位（稍後靠PinkPojo資料對應輸入)
                if( readyDataPojo != null ) stackList.add(readyDataPojo) ;
            }
        }

        // 將 stackList 排序, 部門 > 名稱 > 使用方式 > 日期
        if( CollectionUtils.isNotEmpty( stackList ) ) {
            stackList.sort((o1, o2) -> {
                // 部門 相同
                if (o1.getApartment().compareTo(o2.getApartment()) == 0) {
                    // 名稱相同, 按照使用方式
                    if (o1.getName().compareTo(o2.getName()) == 0) {
                        // 使用方式相同, 按日期決定
                        if (o1.getRestOrMoney().compareTo(o2.getRestOrMoney()) == 0) {
                            return o1.getStartDay().compareTo(o2.getStartDay());
                        }

                        return o1.getRestOrMoney().compareTo(o2.getRestOrMoney());
                    }

                    return o1.getName().compareTo(o2.getName());
                }
                // 部門不相同
                return o1.getApartment().compareTo(o2.getApartment());
            });
        }else
            throw new StopProgramException( wordsPath + "底下沒有成功取得至少一筆Word資料! 請確認該路徑底下有加班單資訊的word文件們.") ;

        return stackList ;
    }
}
