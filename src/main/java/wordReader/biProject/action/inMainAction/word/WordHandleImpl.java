package wordReader.biProject.action.inMainAction.word;

import org.apache.commons.collections4.CollectionUtils;
import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.fileNameFilter.WordsFileFilter;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.util.PropsHandler;
import java.io.File;
import java.io.IOException;
import java.util.*;

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
        // 取得 word存放路徑 底下的以.docx結尾之檔案
        List<File> files = getDocxFileUnderPath(wordsPath);

        if( CollectionUtils.isEmpty(files) )
            throw new StopProgramException("現在" + wordsPath + "底下沒有Word文件") ;

        for( File aFile: files ) {
            // 取得 word 裡面資料 並處理一些計算的部分
            DataPojo readyDataPojo = WordReader.readWord2007Docx(wordsPath+aFile.getName()) ;
            // 如果不為null加到stackList
            if( readyDataPojo != null ) stackList.add(readyDataPojo) ;
        }

        // 將 stackList 排序, 部門 > 名稱 > 使用方式 > 日期
        if( CollectionUtils.isNotEmpty( stackList ) ) {
            stackList.sort(Comparator.comparing(DataPojo::getApartment)
                                     .thenComparing(DataPojo::getName)
                                     .thenComparing(DataPojo::getRestOrMoney)
                                     .thenComparing(DataPojo::getStartDay));
        }else
            throw new StopProgramException( wordsPath + "底下沒有成功取得至少一筆Word資料! 請確認該路徑底下有加班單資訊的word文件們.") ;

        return stackList ;
    }


    /**
     * Return files under @path, which are docx format and not start with '~' and '.'
     *
     * @param path
     * @return File array
     */
    @Override
    public List<File> getDocxFileUnderPath(String path) {
        // 文件過濾器(只接受 .docx結尾並且不是'~'和'.'開頭的文件)
        WordsFileFilter fileFilter = new WordsFileFilter() ;
        // 取得該路徑資料夾
        File dir = new File(path) ;
        // 用文件過濾器取出符合的文件
        File[] validFiles = dir.listFiles(fileFilter);
        // return
        return validFiles != null ? Arrays.asList(validFiles) : null ;
    }

}
