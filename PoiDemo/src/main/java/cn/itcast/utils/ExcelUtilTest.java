package cn.itcast.utils;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import cn.itcast.domain.User;

public class ExcelUtilTest {

	public static void main(String[] args) throws InvalidFormatException, IOException {

		File file = new File("d:/poi.xlsx");

		ExceiUtils<User> eu = new ExceiUtils<User>();

		List<User> users = eu.getEntity(file, new ExcelRowResultHandler<User>() {

			public User invoke(List<Object> list) {
				// list代表的是每一行的数据，我们要将它封装到User对象中。
				User user = new User();
				user.setName((String) list.get(0));
				user.setAge(((Double) list.get(1)).intValue());
				user.setSex((String) list.get(2));
				return user;
			}
		});

		System.out.println(users);
	}
}
