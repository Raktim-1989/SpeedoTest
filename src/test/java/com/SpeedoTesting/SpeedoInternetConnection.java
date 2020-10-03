package com.SpeedoTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class SpeedoInternetConnection {
	
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
	static WebDriver driver;
	
    /*-------------------------------------------------------------------------------------
    |  Method [getExcel]
    |
    |  Purpose:  [This method will create a new Excel spreadsheet in developer's
    |             working directory and will hold the excel object throughout the program.]
    |
    |  Pre-condition:  [APACHE-POI jars or APIS need to be installed for this reporting
    |                    All the xlsx class files implementations are there in POI jar files].
    |
    |  Post-condition: [Once executing this method from java run time enginee a new excel blank
    |                   spreadsheet will be getting generated and fileoutputstream object(fos) 
    |                   will be holding throughout the execution.]
    |  Parameters:
    |      parameter_name -- [Explanation of the purpose of this
    |          parameter to the method.  Write one explanation for each
    |          formal parameter of this method.]
    |
    |  Returns:  [If this method sends back a value via the return
    |      mechanism, describe the purpose of that value here, otherwise
    |      state 'None.' In this case 'None']
    *-------------------------------------------------------------------------------------*/

public static void getExcel()
{
	String dir = System.getProperty("user.dir"); //user.dir 
	System.out.println(dir);
	String path = dir + File.separator + "SpeedoTest.xlsx";
	wb = new XSSFWorkbook(); 
	sh = wb.createSheet("iSpeed");
	
		try {
			fos = new FileOutputStream(path);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	

public static void getCurrentTime()
{
	timestamp = new Timestamp(System.currentTimeMillis()).toString();
	timestamp = timestamp.substring(0, timestamp.length() - 6).replaceAll(":", "");
	filestamp = "./" + "\\Speedo" + " " + timestamp;
	System.out.println(filestamp);
}

	public static WebDriver getDriver() {
		System.setProperty("webdriver.chrome.driver", "Drivers/chromedriver.exe");
		 driver = new ChromeDriver();
		driver.get("https://www.speedtest.net/");
		driver.manage().window().maximize();
		return driver;
	}

/*-------------------------------------------------------------------------------------
|  Method [getHeaderName]
|
|  Purpose:  [This method will create a new row along with three columns respectively.]
|
|  Pre-condition:  [APACHE-POI jars or APIS need to be installed for this reporting
|                    All the xlsx class files implementations are there in POI jar files].
|
|  Post-condition: [Once executing this method from java run time enginee a new excel blank
|                   spreadsheet will be getting generated and fileoutputstream object(fos) 
|                   will be holding throughout the execution.]
|  Parameters:
|      parameter_name -- [Explanation of the purpose of this
|          parameter to the method.  Write one explanation for each
|          formal parameter of this method.]
|
|  Returns:  [If this method sends back a value via the return
|      mechanism, describe the purpose of that value here, otherwise
|      state 'None.' In this case 'None']
*-------------------------------------------------------------------------------------*/


public static void getHeadername()
{
	row = sh.createRow(0);
	Cell cellheaderOne = row.createCell(0);
	cellheaderOne.setCellValue("PING");
	Cell cellheaderTwo = row.createCell(1);
	cellheaderTwo.setCellValue("DOWNLOAD_SPEED");
	Cell cellheaderThree = row.createCell(2);
	cellheaderThree.setCellValue("UPLOAD_SPEED");
	Cell cellheaderFour = row.createCell(3);
	cellheaderFour.setCellValue("TIME");
	Cell cellheaderFive = row.createCell(4);
	cellheaderFive.setCellValue("IP-ADDRESS");
	Cell cellheaderSix = row.createCell(5);
	cellheaderSix.setCellValue("PROVIDER-NAME");
	
	CellStyle style1 = wb.createCellStyle();
	 style1.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
	 style1.setFillPattern(FillPatternType.BIG_SPOTS);
	 //fontcolor
	 Font font = wb.createFont();
   font.setColor(IndexedColors.WHITE.getIndex());
   style1.setFont(font);
   cellheaderOne.setCellStyle(style1);
   cellheaderTwo.setCellStyle(style1);
   cellheaderThree.setCellStyle(style1);
   cellheaderFour.setCellStyle(style1);
   cellheaderFive.setCellStyle(style1);
   cellheaderSix.setCellStyle(style1);
   
    cellStyle = wb.createCellStyle();
	CreationHelper createHelper = wb.getCreationHelper();
	cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("d/m/yy h:mm"));
	LocalDateTime startTime = LocalDateTime.now();
	System.out.println("Start time" + " " + startTime);

   //cell5.setCellStyle(style1);
   //cell6.setCellStyle(style1);
	
}
/*-------------------------------------------------------------------------------------
|  Method [main method-Java Run Time Enginee]
|
|  Purpose:  [The purpose of this method is to extract the live data from the portal and prepare
               a consolidated report for that .]
|
|  Pre-condition:  [APACHE-POI jars or APIS need to be installed for this reporting
|                    All the xlsx class files implementations are there in POI jar files].
|
|  Post-condition: [Once executing this method from java run time enginee a new excel blank
|                   spreadsheet will be getting generated and fileoutputstream object(fos) 
|                   will be holding throughout the execution.]
|  Parameters:
|      parameter_name -- [Explanation of the purpose of this
|          parameter to the method.  Write one explanation for each
|          formal parameter of this method.]
|
|  Returns:  [If this method sends back a value via the return
|      mechanism, describe the purpose of that value here, otherwise
|      state 'None.' In this case 'None']
*-------------------------------------------------------------------------------------*/

	public static void createDataReportMaster() throws InterruptedException, IOException {
		try
		{
		getExcel();	
		getHeadername();
		
		driver.findElement(By.xpath("//span[contains(text(),'Go')]")).click();
		Thread.sleep(50000);
		int counter = 0;
		while(counter <3)
		{
		
		String name = driver.findElement(By.xpath("//*[@class = 'result-container-data']/div[1]/div/div[2]/span")).getText();
		System.out.println(name);
		String download = driver.findElement(By.xpath("//*[@class = 'result-container-data']/div[2]/div/div[2]/span")).getText();
		System.out.println(download);
		WebDriverWait wait = new WebDriverWait(driver, 20);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class = 'result-container-data']/div[3]/div/div[2]/span")));
		String upload = driver.findElement(By.xpath(	"//*[@class = 'result-container-data']/div[3]/div/div[2]/span")).getText();
		System.out.println(upload);
		String ip = driver.findElement(By.xpath("//*[contains(@class , 'js-data-ip')]")).getText();
		String providername = driver.findElement(By.xpath("//*[contains(@class , 'js-data-isp')]")).getText();
	     row = sh.createRow(rownumber++);
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
	
	public static void getDataReportAppend(String filepath) throws IOException
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
	
	
	
////////////////////Main Method/////////////////////////
	public static void main(String[] args) throws InterruptedException, IOException
	{
		getDriver();
		getCurrentTime();
		createDataReportMaster();
		//getDataReportAppend(timestamp);
				
	}
	
	
	
	
	
}
