package top.onceio.excel.utils;

import org.junit.Test;

import java.math.BigDecimal;
import java.util.*;

public class ExcelTplWriterTest {

	@Test
	public void export() {

		List<UserInfo> data = new ArrayList<>();

		for(int i = 0; i <  3; i++) {
			UserInfo ui = new UserInfo();

			ui.setName("name:" + i);
			ui.setBirthday(new Date(System.currentTimeMillis() - i * 365 *24 *60*60000));
			if(i%2==0) {
				ui.setGender("男");
			}else {
				ui.setGender("女");
			}
			ui.setSalary(new BigDecimal((i+1) * 5000));
			data.add(ui);
		}

		Map<String,String> alias = new HashMap<>();
		alias.put("姓名","name");
		alias.put("生日","birthday");
		alias.put("性别","gender");
		alias.put("薪水","salary");

		ExcelClassHelper.write(UserInfo.class,data,alias,"src/test/resources/class-tpl.xlsx","out-class-tpl.xlsx");
	}


}


class UserInfo {
	private String name;
	private String gender;
	private Date birthday;
	private BigDecimal salary;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getGender() {
		return gender;
	}

	public void setGender(String gender) {
		this.gender = gender;
	}

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	public BigDecimal getSalary() {
		return salary;
	}

	public void setSalary(BigDecimal salary) {
		this.salary = salary;
	}
}