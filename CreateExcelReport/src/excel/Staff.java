package excel;

public class Staff {

	private String first_name;

	private String last_name;
	private String staff_number;
	private String phone;
	private String email;

	public Staff(String first_name, String last_name, String staff_number, String phone, String email) {
		super();
		this.first_name = first_name;
		this.last_name = last_name;
		this.staff_number = staff_number;
		this.phone = phone;
		this.email = email;
	}

	public String getFirst_name() {
		return first_name;
	}

	public void setFirst_name(String first_name) {
		this.first_name = first_name;
	}

	public String getLast_name() {
		return last_name;
	}

	public void setLast_name(String last_name) {
		this.last_name = last_name;
	}

	public String getStaff_number() {
		return staff_number;
	}

	public void setStaff_number(String staff_number) {
		this.staff_number = staff_number;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

}
