package wordReader.biProject.action.beforeMainAction;

import org.apache.commons.configuration.ConfigurationException;

import java.io.IOException;

public interface BeforeMain {

    /**
     * 印出歡迎訊息＋相關參數 讓使用者確定是否是正確參數
     */
    void helloMsg() throws IOException;

    /**
     * 如果參數不正確，使用者可以選擇修改任何一個參數設定
     */
    void changeProperties() throws IOException, ConfigurationException;

}
