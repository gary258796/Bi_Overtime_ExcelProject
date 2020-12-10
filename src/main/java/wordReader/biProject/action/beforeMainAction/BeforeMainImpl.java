package wordReader.biProject.action.beforeMainAction;

import org.apache.commons.configuration.ConfigurationException;
import wordReader.biProject.util.PropsHandler;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class BeforeMainImpl implements BeforeMain{

    /**
     * 印出歡迎訊息＋相關參數 讓使用者確定是否是正確參數
     */
    public void helloMsg() throws IOException {
        System.out.println( "**********************************************************************************");
        System.out.println( "Hi! " + PropsHandler.getter("userName") );
        System.out.println( "信箱 : " + PropsHandler.getter("emailAccount") );
        System.out.println( "密碼 : " + PropsHandler.getter("emailPassWord") );
        System.out.println( "加班單word存放路徑  : " + PropsHandler.getter("wordsPath") );
        System.out.println( "產出Excel會在 : " + PropsHandler.getter("writePath")  + "底下");
        System.out.println( "無法處理的特殊(機車)Word會存到 : " + PropsHandler.getter("errorWordsPath")  + "底下");
        System.out.println( "通訊錄Excel(如果加班單沒有截圖,尋找那個人的Email用) : " + PropsHandler.getter("contactPath")  + "底下");
        System.out.println( "目前設定之粉紅顏色對應號碼: " + PropsHandler.getter("pinkValue"));
        System.out.println( "**********************************************************************************");
        System.out.println();
    }

    /**
     * 如果參數不正確，使用者可以選擇修改任何一個參數設定
     */
    public void changeProperties() throws IOException, ConfigurationException {

        BufferedReader consoleInput=new BufferedReader(new InputStreamReader(System.in));

        boolean change = false ;
        char yesNo ;
        while (true) {
            System.out.println("請問有需要更改以上任一資訊嗎？(Y/N)");
            yesNo = consoleInput.readLine().charAt(0);

            if (yesNo == 'Y') {
                change = true;
                break;
            } else if (yesNo == 'N') {
                break;
            } else
                System.out.println("\n 你是故意的還是在靠北我？");
        }

        int chooseNum ;
        while( change ) {
            System.out.println("需要修改哪個？");
            System.out.println("1. 信箱&密碼 2. 加班單Word存放路徑 3. 產出Excel路徑 "
                    + "4. 無法處理Word存放路徑 5. 通訊錄Excel路徑 6. 粉紅顏色(你應該不會用到,是給專業的屎用的) "
                    + "7. 查看現在狀態 8. 結束");

            chooseNum = Integer.parseInt(consoleInput.readLine()) ;
            String inputString;
            switch (chooseNum) {

                case 1:
                    System.out.print("請輸入信箱(輸入完成請按下Enter, 不更改直接按下Enter) : ");
                    String emailString = consoleInput.readLine();
                    System.out.print("請輸入密碼(輸入完成請按下Enter, 不更改直接按下Enter) : ");
                    String passWordString = consoleInput.readLine();
                    PropsHandler.setter("emailAccount", emailString);
                    PropsHandler.setter("emailPassWord", passWordString);
                    break;
                case 2:
                    System.out.print("請輸入new加班單Word存放路徑(輸入完成請按下Enter, 不更改直接按下Enter): ");
                    inputString = consoleInput.readLine();
                    PropsHandler.setter("wordsPath", inputString);
                    break;
                case 3:
                    System.out.print("請輸入new產出Excel存放路徑(輸入完成請按下Enter, 不更改直接按下Enter): ");
                    inputString = consoleInput.readLine();
                    PropsHandler.setter("writePath", inputString);
                    break;
                case 4:
                    System.out.print("請輸入new無法處理Word存放路徑(輸入完成請按下Enter, 不更改直接按下Enter): ");
                    inputString = consoleInput.readLine();
                    PropsHandler.setter("errorWordsPath", inputString);
                    break;
                case 5:
                    System.out.print("請輸入new通訊錄Excel路徑(輸入完成請按下Enter, 不更改直接按下Enter): ");
                    inputString = consoleInput.readLine();
                    PropsHandler.setter("contactPath", inputString);
                    break;
                case 6:
                    System.out.print("請輸入new Pink Index(你調皮偷亂改後果很嚴重喔,輸入完成請按下Enter, 不更改直接按下Enter): ");
                    int inputInt = Integer.parseInt(consoleInput.readLine());
                    PropsHandler.setter("pinkValue", Integer.toString(inputInt));
                    break;
                case 7:
                    helloMsg();
                    break;
                case 8:
                    change=false;
                    break;
                default:
                    System.out.println("問號,你是有看到你選的這個選項??");

            }
        }

        System.out.println("\n 產出Excel....");
    }
}
