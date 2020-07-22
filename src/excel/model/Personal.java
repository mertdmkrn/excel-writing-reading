package excel.model;

public class Personal {
	
	private int perId;
	private String perName;
	private String perDept;
	private int perSalary;
	
	public Personal(int perId, String perName, String perDept, int perSalary) {
		this.perId = perId;
		this.perName = perName;
		this.perDept = perDept;
		this.perSalary = perSalary;
	}

	public int getPerId() {
		return perId;
	}

	public void setPerId(int perId) {
		this.perId = perId;
	}

	public String getPerName() {
		return perName;
	}

	public void setPerName(String perName) {
		this.perName = perName;
	}

	public String getPerDept() {
		return perDept;
	}

	public void setPerDept(String perDept) {
		this.perDept = perDept;
	}

	public int getPerSalary() {
		return perSalary;
	}

	public void setPerSalary(int perSalary) {
		this.perSalary = perSalary;
	}
	
	
	
	

}
