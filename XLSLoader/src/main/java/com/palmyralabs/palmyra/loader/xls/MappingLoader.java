package com.palmyralabs.palmyra.loader.xls;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class MappingLoader {

	private String mappingFile; // "defect_mapping.properties"
	private List<Mapping> mappings;

	public MappingLoader(String mappingFile){
		this.mappingFile = mappingFile;
		try {
			this.loadMapping();
		} catch (IOException e) {
			throw new RuntimeException("Mapping file " + mappingFile + " cannot be loaded", e);
		}
	}

	public void loadMapping() throws IOException {
		InputStream inStream = MappingLoader.class.getClassLoader().getResourceAsStream(this.mappingFile);
		Properties props = new Properties();
		props.load(inStream);
		mappings = getMapping(props);
	}

	public List<Mapping> getMappings() {
		return mappings;
	}

	private List<Mapping> getMapping(Properties props) {
		List<Mapping> result = new ArrayList<Mapping>();
		String index = null;
		String dataType = null;
		String defValue = null;
		for (String key : props.stringPropertyNames()) {
			if (key.contains(".")) {
				int pos = key.indexOf(".");
				index = key.substring(0, pos);
				dataType = key.substring(pos + 1, key.length());
				defValue = null;
			} else if (key.contains("_")) {
				int pos = key.indexOf("_");
				index = key.substring(0, pos);
				defValue = key.substring(pos + 1, key.length());
				dataType = null;
			}			
			else{
				index = key;
				dataType = null;
				defValue = null;
			}
			try {
				int position = Integer.parseInt(index) - 1;
				String value = props.getProperty(key);
				if (null != value && value.length() > 1) {
					if (value.endsWith("!")) {
						value = value.substring(0, value.length() -1);
						Mapping mapping = new Mapping(position, dataType, value, true);
						result.add(mapping);
						mapping.setStaticValue(defValue);
					} else {
						Mapping mapping = new Mapping(position, dataType, value, false);
						mapping.setStaticValue(defValue);
						result.add(mapping);
					}
				}
			} catch (NumberFormatException nfe) {
				continue;
			}
		}

		return result;
	}
}
