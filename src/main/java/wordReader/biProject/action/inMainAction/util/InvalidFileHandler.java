package wordReader.biProject.action.inMainAction.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;

/**
 * 處理錯誤的Word加班單檔案
 */
public class InvalidFileHandler {

    /**
     * 有問題檔案，丟到設定路徑底下
     * @param filePath     : 錯誤文件的路徑
     * @param destination  : 錯誤文件所要放置的路徑
     * @throws IOException
     */
    public static void throwInvalidFilePathFileToDestination(String filePath, String destination) throws IOException {
        // 錯誤文件的檔案名稱
        String fileName = filePath.substring(filePath.lastIndexOf("/"));
        // 將錯誤文件輸出到路徑
        FileChannel in = new FileInputStream( filePath ).getChannel();
        FileChannel out = new FileOutputStream( destination+fileName ).getChannel();
        out.transferFrom( in, 0, in.size() );
        in.close();
        out.close();

        System.out.println(fileName + " 有問題,已加入到指定error word路徑！");
    }

}
