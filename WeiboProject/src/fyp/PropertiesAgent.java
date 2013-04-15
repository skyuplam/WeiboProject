package fyp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

public class PropertiesAgent { 
	private String propertiesFileName = "weibo_config.properties";
	private String propertiesFileComment = "";
	private Properties prop = null;
	
	public PropertiesAgent(){
		this.prop = new Properties();
	}
	
	public PropertiesAgent(String propertiesFileName, String propertiesFileComment){
		this.propertiesFileName = propertiesFileName;
		this.propertiesFileComment = propertiesFileComment;
		this.prop = new Properties();
	}
	
	private void initProperties(){
		if(prop == null){
			prop = new Properties();
		}
	}
	
	public void writeProperty(String key, String value){
		initProperties();
		prop.setProperty(key, value);
		storeProperties();
	}

	private void storeProperties() {
		try {
			prop.store(new FileOutputStream(propertiesFileName), propertiesFileComment);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void writeProperties(Map<String, String> propertiesMap){
		initProperties();
		Set<String> keyset = propertiesMap.keySet();
		Iterator<String> iKeyset = keyset.iterator();
		while(iKeyset.hasNext()){
			String key = iKeyset.next();
			prop.setProperty(key, propertiesMap.get(key));
		}
		storeProperties();
	}
	
	public String loadProperty(String key, String defaultValue){
		initProperties();
		String value = null;
		try {
			File file = new File(propertiesFileName);
			if(!file.exists()){
				return null;
			}
			FileInputStream fileIS = new FileInputStream(file);
			prop.load(fileIS);
			value = prop.getProperty(key, defaultValue);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return value;
	}
	
	public Map<String, String> getAllProperties(){
		Map<String, String> propertiesMap = new HashMap<String, String>();
		try {
			File file = new File(propertiesFileName);
			if(!file.exists()){
				return null;
			}
			FileInputStream fileIS = new FileInputStream(file);
			prop.load(fileIS);
			Set<Object> keyset = prop.keySet();
			Iterator<Object> iKeyset = keyset.iterator();
			while(iKeyset.hasNext()){
				String key = (String)iKeyset.next();
				propertiesMap.put(key, prop.getProperty(key));
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return propertiesMap;
	}
}
