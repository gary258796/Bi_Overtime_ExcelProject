package wordReader.biProject.model;

import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 從震旦雲抓取的資料
 */
@Data
@NoArgsConstructor
public class PinkPojo {

	private String date ;
	private String employee ;
	private String onTime ;
	private String offTime ;
	private String missContent ;
	// 方便計算用
	private int startHour ;
	private int startMin ;
	private int endHour ;
	private int endMin ;

}
