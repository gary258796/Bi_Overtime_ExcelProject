package wordReader.biProject.model;

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
	
	
	public PinkPojo() {}

	public PinkPojo(String date, String employee, String onTime, String offTime, String missContent, int startHour,
			int startMin, int endHour, int endMin) {
		super();
		this.date = date;
		this.employee = employee;
		this.onTime = onTime;
		this.offTime = offTime;
		this.missContent = missContent;
		this.startHour = startHour;
		this.startMin = startMin;
		this.endHour = endHour;
		this.endMin = endMin;
	}



	// Getter and Setter

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	public String getEmployee() {
		return employee;
	}

	public void setEmployee(String employee) {
		this.employee = employee;
	}

	public String getOnTime() {
		return onTime;
	}

	public void setOnTime(String onTime) {
		this.onTime = onTime;
	}

	public String getOffTime() {
		return offTime;
	}

	public void setOffTime(String offTime) {
		this.offTime = offTime;
	}

	public String getMissContent() {
		return missContent;
	}

	public void setMissContent(String missContent) {
		this.missContent = missContent;
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
	
	// Getter and Setter
	
	
}
