package com.SpeedoTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class SpeedoMasterSheet {
	public static WebDriver driver;
	static XSSFWorkbook wb ;
	static XSSFSheet sh;
	static FileOutputStream fos;
	static File file;
	static XSSFRow row;
	static XSSFCell cell;
	static int rownumber = 1;
	static int cellnumber = 0;
	static CellStyle cellStyle;
	static String timestamp;
	static String filestamp;
	static int lastrow;
	
	
	public static void getDataReportAppend() throws IOException
	{
		try
		{
		//String path = filepath + ".xlsx" ;	
			String path = System.getProperty("user.dir") + File.separator + "SpeedoTest.xlsx"
;		FileInputStream fis = new FileInputStream(path);
		wb = new XSSFWorkbook(fis);
		System.out.println("done");
		sh = wb.getSheetAt(0);
	    lastrow = sh.getLastRowNum();
	    fos = new FileOutputStream(path);
	    
	    driver.findElement(By.xpath("//span[contains(text(),'Go')]")).click();
		Thread.sleep(50000);
		int counter = 0;
		while(counter <5)
		{
		
		String name = driver.findElement(By.xpath("//*[@class = 'result-container-data']/div[1]/div/div[2]/span")).getText();
		System.out.println(name);
		String download = driver.findElement(By.xpath("//*[@class = 'result-container-data']/div[2]/div/div[2]/span")).getText();
		System.out.println(download);
		WebDriverWait wait = new WebDriverWait(driver, 20);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class = 'result-container-data']/div[3]/div/div[2]/span")));
		String upload = driver.findElement(By.xpath("//*[@class = 'result-container-data']/div[3]/div/div[2]/span")).getText();
		System.out.println(upload);
		String ip = driver.findElement(By.xpath("//*[contains(@class , 'js-data-ip')]")).getText();
		String providername = driver.findElement(By.xpath("//*[contains(@class , 'js-data-isp')]")).getText();
	     row = sh.createRow(++lastrow);
	     cellStyle = wb.createCellStyle();
	 	CreationHelper createHelper = wb.getCreationHelper();
	 	cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("d/m/yy h:mm"));
	 	LocalDateTime startTime = LocalDateTime.now();
	 	System.out.println("Start time" + " " + startTime);
	    Cell cellPng  = row.createCell(0);
	    Cell cellDownload  = row.createCell(1);
	    Cell cellUpload  = row.createCell(2);
	    Cell time = row.createCell(3);
	    Cell cellIp = row.createCell(4);
	    Cell cellProvider = row.createCell(5);
	    time.setCellStyle(cellStyle);	
	    Date current = new Date();
	    time.setCellValue(new Date());
	     cellPng.setCellValue(name);  
	     cellDownload.setCellValue(download	);
	     cellUpload.setCellValue(upload);
	     cellIp.setCellValue(ip);
	     cellProvider.setCellValue(providername);
		System.out.println(name +  " " + upload + " " + download + " " + current);
		driver.findElement(By.xpath("//span[contains(text(),'Go')][@class = 'start-text']")).click();
		//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("")));
		Thread.sleep(50000);	
		counter ++;
		}
		
		wb.write(fos);
		fos.close();

	}
	
	catch(Exception e)
	{
		wb.write(fos);
		fos.close();
	
	}
		
	}
	


	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		driver = SpeedoInternetConnection.getDriver();
		getDataReportAppend();
		

	}

}
