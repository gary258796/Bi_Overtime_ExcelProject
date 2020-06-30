package wordReader.biProject.model;

public class DataPojo {

	private String apartment ; 
	private String name ; 
	private String startDay ; 
	private String startTime ; 
	private String endDay ; 
	private String endTime ; 
	private String date ; // 申請日
	private String applyHour ; 
	private String projectName ; //
	private String reason ; 
	private String admitTime; 
	private String restOrMoney ; // 使用方式 
	private String extraMsg ; 
	// 
	private String actualStartTime ;
	private String actualEndTime; 
	private int differTotalTime ; // 對照申請時數
	private boolean isSunday ;
	private String missContent ;
	// 
	private boolean hasPhoto ; 
	
	
	// 方便計算用的欄位
	private int startHour ;
	private int startMin ;
	private int endHour ;
	private int endMin ;
	
	// 額外缺
	
	public DataPojo() {
		super();
	}

	public DataPojo(String apartment, String name, String startDay, String startTime, String endDay, String endTime,
			String date, String applyHour, String projectName, String reason, String admitTime, String restOrMoney,
			String extraMsg, String actualStartTime, String actualEndTime, int differTotalTime, boolean isSunday,
			String missContent, boolean hasPhoto) {
		super();
		this.apartment = apartment;
		this.name = name;
		this.startDay = startDay;
		this.startTime = startTime;
		this.endDay = endDay;
		this.endTime = endTime;
		this.date = date;
		this.applyHour = applyHour;
		this.projectName = projectName;
		this.reason = reason;
		this.admitTime = admitTime;
		this.restOrMoney = restOrMoney;
		this.extraMsg = extraMsg;
		this.actualStartTime = actualStartTime;
		this.actualEndTime = actualEndTime;
		this.differTotalTime = differTotalTime;
		this.isSunday = isSunday;
		this.missContent = missContent;
		this.hasPhoto = hasPhoto ; 
	}
	
	

	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getReason() {
		return reason;
	}
	public void setReason(String reason) {
		this.reason = reason;
	}
	public String getApartment() {
		return apartment;
	}
	public void setApartment(String apartment) {
		this.apartment = apartment;
	}
	public String getStartTime() {
		return startTime;
	}
	public void setStartTime(String startTime) {
		this.startTime = startTime;
	}
	public String getEndTime() {
		return endTime;
	}
	public void setEndTime(String endTime) {
		this.endTime = endTime;
	}

	public String getRestOrMoney() {
		return restOrMoney;
	}

	public void setRestOrMoney(String restOrMoney) {
		this.restOrMoney = restOrMoney;
	}

	public String getStartDay() {
		return startDay;
	}

	public void setStartDay(String startDay) {
		this.startDay = startDay;
	}

	public String getEndDay() {
		return endDay;
	}

	public void setEndDay(String endDay) {
		this.endDay = endDay;
	}

	public String getApplyHour() {
		return applyHour;
	}

	public void setApplyHour(String applyHour) {
		this.applyHour = applyHour;
	}

	public String getProjectName() {
		return projectName;
	}

	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}

	public String getAdmitTime() {
		return admitTime;
	}

	public void setAdmitTime(String admitTime) {
		this.admitTime = admitTime;
	}

	public String getActualStartTime() {
		return actualStartTime;
	}

	public void setActualStartTime(String actualStartTime) {
		this.actualStartTime = actualStartTime;
	}

	public String getActualEndTime() {
		return actualEndTime;
	}

	public void setActualEndTime(String actualEndTime) {
		this.actualEndTime = actualEndTime;
	}

	public int getDifferTotalTime() {
		return differTotalTime;
	}

	public void setDifferTotalTime(int differTotalTime) {
		this.differTotalTime = differTotalTime;
	}

	public boolean isSunday() {
		return isSunday;
	}

	public void setSunday(boolean isSunday) {
		this.isSunday = isSunday;
	}

	public int getStartHour() {
		return startHour;
	}

	public void setStartHour(int startHour) {
		this.startHour = startHour;
	}

	public int getStartMin() {
		return startMin;
	}

	public void setStartMin(int startMin) {
		this.startMin = startMin;
	}

	public int getEndHour() {
		return endHour;
	}

	public void setEndHour(int endHour) {
		this.endHour = endHour;
	}

	public int getEndMin() {
		return endMin;
	}

	public void setEndMin(int endMin) {
		this.endMin = endMin;
	}

	public String getMissContent() {
		return missContent;
	}

	public void setMissContent(String missContent) {
		this.missContent = missContent;
	}

	public String getExtraMsg() {
		return extraMsg;
	}

	public void setExtraMsg(String extraMsg) {
		this.extraMsg = extraMsg;
	}

	public boolean isHasPhoto() {
		return hasPhoto;
	}

	public void setHasPhoto(boolean hasPhoto) {
		this.hasPhoto = hasPhoto;
	} 

	
}
