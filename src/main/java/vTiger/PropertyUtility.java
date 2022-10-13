package vTiger;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class PropertyUtility {

	public String readDataFromPropertyFile(String key) throws IOException {
		FileInputStream fis=new FileInputStream(ConstantUtility.propertyPath);
		Properties pobj=new Properties();
		pobj.load(fis);
		String value=pobj.getProperty(key);
		return value;
	}
}
