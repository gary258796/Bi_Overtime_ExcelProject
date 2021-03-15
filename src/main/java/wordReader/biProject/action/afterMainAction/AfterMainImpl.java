package wordReader.biProject.action.afterMainAction;

import com.google.common.collect.Iterables;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Workbook;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.model.PinkPojo;
import wordReader.biProject.util.PropsHandler;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class AfterMainImpl implements AfterMain{

    /**
     * 用取好之資料，產出Excel
     * @param pinkPojos :
     * @param dataPojos :
     * @throws ParseException
     */
    @Override
    public void finalProcess(List<PinkPojo> pinkPojos , List<DataPojo> dataPojos) throws ParseException {

        List<PinkPojo> removePinks = new ArrayList<>() ;

        // 每一筆dataPojo(Word加班單), 尋找(By名字、日期)有沒有對應的pinkPojo(震旦雲六日資料)
        for( DataPojo data: dataPojos ) {
            for( PinkPojo pinkPojo : pinkPojos ) {
                // 有的話計算相關資料並且更新dataPojo對應欄位, 並且把該pinkPojo從pinkPoJos移除
                if(  CalculateTool.pinkNameHandle( pinkPojo.getEmployee() ).equals(data.getName())
                        && CalculateTool.pinkDateHandle( pinkPojo.getDate() ).equals(data.getStartDay()) ) {
                    // 實際上班時間
                    data.setActualStartTime( pinkPojo.getOnTime() );
                    // 實際下班時間
                    data.setActualEndTime( pinkPojo.getOffTime() );
                    // 是否需調整( 是否為星期天)
                    data.setSunday(CalculateTool.isSunday(data.getStartDay())) ;
                    // 缺卡內容
                    data.setMissContent( pinkPojo.getMissContent() );
                    // 對照申請時數 ( 實際上-實際下 ) - (申請上-申請下) , 存分鐘
                    if(!pinkPojo.getOnTime().equals("") && !pinkPojo.getOffTime().equals("")) {
                        data.setDifferTotalTime( CalculateTool.calDifferTime(pinkPojo.getStartHour(), pinkPojo.getStartMin(), pinkPojo.getEndHour(), pinkPojo.getEndMin())
                                - CalculateTool.calDifferTime(data.getStartHour(), data.getStartMin(), data.getEndHour(), data.getEndMin()) );
                    }else
                        data.setDifferTotalTime( 0 );

                    removePinks.add(pinkPojo) ;
                }
            }
        }

        for( PinkPojo removedPinkPojo : removePinks) pinkPojos.remove(removedPinkPojo) ;
        // 保險起見清空list裡面null，不管有沒有
        if( CollectionUtils.isNotEmpty(dataPojos))
            Iterables.removeIf(dataPojos, Objects::isNull);
        if( CollectionUtils.isNotEmpty(pinkPojos))
            Iterables.removeIf(pinkPojos, Objects::isNull);

        // Write excel
        try {
            Workbook workbook = ExcelWriter.exportData(dataPojos, pinkPojos) ; // POI會幫我們處理所有格式上所需
            FileOutputStream out=new FileOutputStream(PropsHandler.getter("writePath"));
            workbook.write(out);
            System.out.println("建立Excel成功\n");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 寄信給忘記截圖的同仁
     * @param dataPoJos :
     * @throws IOException :
     */
    public void sendMail(List<DataPojo> dataPoJos) throws IOException {

        BufferedReader consoleInput=new BufferedReader(new InputStreamReader(System.in));
        System.out.println("要寄信給加班單沒有截圖的人嗎？(Y/N)");
        char yesNo ;
        yesNo = consoleInput.readLine().charAt(0);

        if( yesNo == 'Y' ){
            System.out.println("\n寄信中....");
            for( DataPojo d : dataPoJos ) {
                if( !d.isHasPhoto() ) {
                    // send mail
                    OutlookSender.sendMail(d.getName(), d.getStartDay());
                }
            }

            System.out.println("信件全數送完");
        }
    }

}
