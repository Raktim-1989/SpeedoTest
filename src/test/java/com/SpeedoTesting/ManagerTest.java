package com.SpeedoTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ManagerTest {
	
	static Properties property;
	
	public static Properties setProperties() throws IOException
	{
		File file = new File("./" + "Configuration/config.properties");
		FileInputStream fis = new FileInputStream(file);
		property = new Properties();
		property.load(fis);
		return property;
	}

	
}
