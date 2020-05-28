package wordReader.biProject;

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
	private boolean restOrMoney ;
	
	// 額外缺
	
	public DataPojo() {
		super();
	}
	
	public DataPojo(String apartment, String name, String startDay, String startTime, String endDay, String endTime,
			String date, String applyHour, String projectName, String reason, String admitTime, boolean restOrMoney) {
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
	public boolean isRestOrMoney() {
		return restOrMoney;
	}
	public void setRestOrMoney(boolean restOrMoney) {
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
		
	
}
