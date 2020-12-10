package wordReader.biProject.action.afterMainAction;

import wordReader.biProject.util.Time;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class CalculateTool {

    /**
     * 轉換從Excel取到的日期格式 yyyy-mm-dd -> yyyy/mm/dd
     * @param pinkDate : 獲取的日期，格式是yyyy-mm-dd
     * @return : yyyy/mm/dd
     */
    public static String pinkDateHandle(String pinkDate) {

        // 轉成 yyyy/mm/dd 形式
        int first = pinkDate.indexOf('-');
        int second = pinkDate.indexOf("-", first + 1);

        String year= pinkDate.substring(0, first ) ;
        String month= pinkDate.substring(first + 1, second ) ;
        String date= pinkDate.substring(second + 1, pinkDate.indexOf('(') ) ;

        return year + "/" + month + "/" + date ;
    }

    /**
     * 從參數裡取出名稱回傳
     * @param pinkName: 震旦雲資料取出的資料： 會是部門+姓名
     * @return : 姓名
     */
    public static String pinkNameHandle(String pinkName) {

        String ret_Str = pinkName.substring(pinkName.lastIndexOf('-') + 1) ;

        return cleanString(ret_Str) ;
    }


    /**
     * 看日期是否為星期天
     * @param bDate :
     * @return
     * @throws ParseException
     */
    public static boolean isSunday(String bDate) throws ParseException {

        DateFormat format1 = new SimpleDateFormat("yyyy/MM/dd");
        Date bdate = format1.parse(bDate);
        Calendar cal = Calendar.getInstance();
        cal.setTime(bdate);
        if( cal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
            return true ;
        }

        return false ;
    }

    /**
     * 計算兩個時間的差,回傳分鐘數
     * @param startHour
     * @param startMin
     * @param endHour
     * @param endMin
     * @return
     */
    public static int calDifferTime(int startHour, int startMin, int endHour, int endMin) {

        Time start = new Time(startHour, startMin) ;
        Time end = new Time(endHour, endMin) ;

        Time differTime = Time.diffTime(start, end);

        return differTime.getHours() * 60 + differTime.getMinutes() ;

    }

    private static String cleanString(String unCleanString) {
        return unCleanString.trim().replaceAll("\r\n|\r|\n|\t|\f|\b", "").replaceAll("\\s+","")
                .replaceAll("[　*| *| *|//s*]*", "").replaceAll("_", "") ;
    }

}
