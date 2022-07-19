package com.palmyralabs.palmyra.loader.xls;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Config {
	public static final String CONFIG_FILE = "application.properties";
	
	protected Properties props = new Properties();
	private static Config config = null;
	protected static Logger logger = LoggerFactory.getLogger(Config.class);

	protected Config() throws IOException {
		props.load(getConfigStream());
	}

	protected InputStream getConfigStream() {
		return ClassLoader.getSystemResourceAsStream(CONFIG_FILE);
	}

	public static Config getInstance() {
		if (null != config)
			return config;

		try {
			config = new Config();
		} catch (IOException e) {
			logger.error("Config file " + CONFIG_FILE + " not found", e);
			return null;
		}
		return config;
	}

	public String getServerURL() {
		return props.getProperty("fluwiz.server.url");
	}

	public String getContextPath() {
		return props.getProperty("fluwiz.server.context");
	}

	public String getServerUser() {
		return props.getProperty("fluwiz.server.username");
	}

	public String getServerPassword() {
		return props.getProperty("fluwiz.server.password");
	}

	public String getProperty(String key) {
		return props.getProperty(key);
	}

	public String getProperty(String key, String defaultValue) {
		String value = props.getProperty(key);
		if (null == value)
			return defaultValue;

		return value;
	}

	public long getLongProperty(String key, long defValue) {
		String sValue = props.getProperty(key);
		try {
			return Long.parseLong(sValue);
		} catch (Exception e) {
			return defValue;
		}
	}

	public int getIntProperty(String key, int defValue) {
		String sValue = props.getProperty(key);
		try {
			return Integer.parseInt(sValue);
		} catch (Exception e) {
			return defValue;
		}
	}
}
