package wordReader.biProject.action.inMainAction.word;

import wordReader.biProject.cusError.StopProgramException;
import wordReader.biProject.util.Time;

/**
 * 針對word取得加班單各個欄位資料
 */
public class FieldExtractHandler {

    private final String wordBody;

    public FieldExtractHandler(String wordBody) {
        this.wordBody = wordBody;
    }

    // 申請日期
    public String getApplyDate() throws StopProgramException {
        String startString = "申請日期",endString = "部門名稱" ;
        String dateString =  getFiledCommon(startString, endString);
        String yearString = dateString.substring(0, dateString.indexOf('/')) ;
        String monthString = dateString.substring( dateString.indexOf('/') + 1, dateString.indexOf('/', dateString.indexOf('/') + 1) );
        String dayString = dateString.substring( dateString.lastIndexOf('/') + 1) ;

        // 二位數顯示月份跟日期
        if( Integer.parseInt(monthString) < 10 )
            monthString = "0" + monthString;
        if( Integer.parseInt(dayString) < 10 )
            dayString = "0" + dayString ;

        return yearString + "/" + monthString + "/" + dayString ;
    }

    // 取得部門
    public String getApartment() throws StopProgramException {
        String startString = "部門名稱", endString = "員工姓名" ;
        return getFiledCommon(startString, endString);
    }

    // 取得員工姓名
    public String getStaffName() throws StopProgramException {
        String startString = "員工姓名(中文)", endString = "實際加班日期";
        return getFiledCommon(startString, endString);
    }

    // 取得實際加班日期
    public String getActualDate( ) throws StopProgramException {
        String startString = "實際加班日期", endString = "實際加班時間";
        String dateString =  getFiledCommon(startString, endString);
        String yearString = dateString.substring(0, dateString.indexOf('/')) ;
        String monthString = dateString.substring( dateString.indexOf('/') + 1, dateString.indexOf('/', dateString.indexOf('/') + 1) );
        String dayString = dateString.substring( dateString.lastIndexOf('/')+ 1 ) ;

        // 二位數顯示月份跟日期
        if( Integer.parseInt(monthString) < 10 )
            monthString = "0" + monthString ;

        if( Integer.parseInt(dayString) < 10 )
            dayString = "0" + dayString ;

        return yearString + "/" + monthString + "/" + dayString ;
    }

    // 取得實際加班時間(開始)
    public String getStartTime( ) throws StopProgramException {
        String startString = "(起)", endString = "(迄)";
        String hourAndMinuteString =  getFiledCommon(startString, endString);

        int index = hourAndMinuteString.indexOf("時") ;
        String hourString = cleanString(hourAndMinuteString.substring(0, index)) ;
        String minString = cleanString(hourAndMinuteString.substring(index+1, hourAndMinuteString.indexOf("分") )) ;

        return hourString + ":" + minString  ;
    }

    // 取得實際加班時間(結束)
    public String getEndTime( ) throws StopProgramException {
        String startString = "(迄)", endString = "專案名稱" ;
        String hourAndMinuteString =  getFiledCommon(startString, endString);

        int index = hourAndMinuteString.indexOf("時") ;
        String hourString = cleanString(hourAndMinuteString.substring(0, index)) ;
        String minString = cleanString(hourAndMinuteString.substring(index+1, hourAndMinuteString.indexOf("分") )) ;

        return hourString + ":" + minString  ;
    }

    // 取得 專案名稱
    public String getProjectName( ) throws StopProgramException {
        String startString = "(請填寫完整)", endString = "加班事由";
        return getFiledCommon(startString, endString);
    }

    // 取得 加班事由
    public String getLateReason( ) throws StopProgramException {
        String startString = "加班事由", endString = "使用方式";
        return getFiledCommon(startString, endString);
    }

    // 使用方式
    public String restOrMoney( ) throws StopProgramException {
        String startString = "使用方式", endString = "備註";
        return getFiledCommon(startString, endString);
    }

    // 取得 備註
    public String getExtraMsg( ) throws StopProgramException {
        String startString = "備註", endString = "注意事項";
        return getFiledCommon(startString, endString);
    }

    // 取得起始小時
    public Integer getStartHour() throws StopProgramException {
        String startString = getStartTime();
        String startHourString = startString.substring(0, startString.indexOf(':'));
        return Integer.parseInt(startHourString);
    }

    // 取得起始分鐘
    public Integer getStartMinute() throws StopProgramException {
        String startString = getStartTime();
        String startMinString = startString.substring(startString.indexOf(':') + 1) ;
        return Integer.parseInt(startMinString);
    }

    // 取得結束小時
    public Integer getEndHour() throws StopProgramException  {
        String endString = getEndTime();
        String endHourString = endString.substring( 0, endString.indexOf(':')) ;
        return Integer.parseInt(endHourString);
    }

    // 取得結束分鐘
    public Integer getEndMinute() throws StopProgramException {
        String endString = getEndTime();
        String endMinString = endString.substring(endString.indexOf(':') + 1) ;
        return Integer.parseInt(endMinString);
    }

    // 取得申請時數， = 總共工時(結束時間 - 開始時間)
    public String getApplyHour() throws StopProgramException {
        Time startTime = new Time(getStartHour(), getStartMinute());
        Time endTime = new Time(getEndHour(), getEndMinute());
        Time applyHour = Time.diffTime(startTime, endTime);
        String hourStr = applyHour.getHours() < 10 ? "0" + applyHour.getHours() : Integer.toString(applyHour.getHours());
        String minStr = applyHour.getMinutes() < 10 ? "0" + applyHour.getMinutes() : Integer.toString(applyHour.getMinutes());
        return hourStr + "時:" +  minStr + "分";
    }

    // 取得承認工時 (中午12~1點 以及 晚上6~7點是不算時數的(吃飯時間))
    public String getAdmitTime() throws StopProgramException {
        Time startTime = new Time(getStartHour(), getStartMinute());
        Time endTime = new Time(getEndHour(), getEndMinute());
        Time admitTime = Time.getTotalHourNMins(startTime, endTime);
        // 超過30分鐘 直接進位多一小時
        if( admitTime.getMinutes() >= 30 )
            admitTime.setHours( admitTime.getHours() + 1 );

        return Integer.toString(admitTime.getHours());
    }

    // Ｗord 有無上傳圖片
    public boolean isImageInOrNot( String xmlString ) {
        return xmlString.contains("圖片");
    }

    // 取得欄位的共用部分
    private String getFiledCommon( String startString, String endString) throws StopProgramException {

        int startIndex = wordBody.indexOf(startString) ;
        int endIndex = wordBody.indexOf(endString) ;
        int offset = startString.length() ;

        if(startIndex == -1 || endIndex == -1)
            throw new StopProgramException("Can't correctly Extract fields in word.");

        return cleanString(wordBody.substring(startIndex+offset, endIndex));
    }

    // 去除String裡面空白換行等字元
    private String cleanString(String unCleanString) {
        return unCleanString.trim().replaceAll("\r\n|\r|\n|\t|\f|\b", "") ;
    }

}
