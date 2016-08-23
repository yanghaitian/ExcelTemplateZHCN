package com.yht.exceltemplate;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.yht.exceltemplate.ExcelTemplate;

public class Test {
	public static void main(String[] args) {
		try {
			InputStream is = new FileInputStream("E:/project/myProjects/ExcelTemplateZHCN/src/main/resources/test.xls");
			OutputStream outputStream = new FileOutputStream("E:/b.xls");
			Map<String,Object> map = new HashMap<String, Object>();
			map.put("a", "小a");
			map.put("b", "大b");
			map.put("c", "嘿嘿c");

			List<Map> list1 = new ArrayList<Map>();
			Map<String,Object> list1Map1 = new HashMap<String,Object>();
			list1Map1.put("哈哈", "哈哈嘿嘿");
			list1Map1.put("haha", "hahahehe");//
			list1.add(list1Map1);

			Map<String,Object> list1Map2 = new HashMap<String,Object>();
			list1Map2.put("哈哈", "哈哈嘿嘿2");
			list1Map2.put("haha", "hahahehe2");
			list1.add(list1Map2);
			map.put("list1", list1);

			ExcelTemplate excelTemplate = new ExcelTemplate(map);
			excelTemplate.printTemplate(is,outputStream);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}