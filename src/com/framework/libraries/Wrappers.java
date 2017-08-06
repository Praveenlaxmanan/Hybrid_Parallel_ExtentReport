package com.framework.libraries;
 
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
 
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;
import org.testng.collections.Lists;
 
import com.beust.jcommander.internal.Maps;
import com.framework.libraries.DataDrivenExcelSheet;
import com.relevantcodes.extentreports.LogStatus;
 
public class Wrappers extends BaseTestReport {
               
	public ExcelReader envExcel;
	public ExcelReader suiteExcel;
	public Properties propValue = null;
	public FileInputStream fip;
	public String sheetName = null;
	public WebDriver driver;
	public ExcelReader functionExcel;
	public static Wrappers wrapper = new Wrappers();
               
	//Get the environment details using HashMap
	public Map<String, HashMap<String, String>> getEnvironment(){
		
		//Hashmap parameters - >>Map with Hashmap
		Map<String,HashMap<String,String>> hmap = Maps.newHashMap();
                               
		  try{
                                               
			  for(int i=1; i<envExcel.getNoOfRows(); i++){
				  String statusCol = envExcel.getDataByColumn(i, "Status");
				  
				  if(statusCol.equalsIgnoreCase("Yes")){
					  //Store the environment details in Hashmap
					  HashMap<String,String> map = new HashMap<>();        
					  String browserName = envExcel.getDataByColumn(i, "BrowerName");
					  String browserURL = envExcel.getDataByColumn(i, "URL");
					  String testSuiteName = envExcel.getDataByColumn(i, "TestSuite_Name");
//                    enValues = enValues + browserName +" ## " +browserURL +" $$ "+testSuiteName + " ";
					  
					  map.put("BN", browserName);
					  map.put("Url", browserURL);
					  //Store Hashmap parameters in Hashmap
					  hmap.put(testSuiteName, map);
				  }
                               
			  }
			  return hmap;
		  }catch(Exception e){
                                               
			  System.out.println("SheetName is not available in test environment");
		  }
		  return null;
	}
               
	public void setSheet(String sheetName) throws FileNotFoundException, IOException{
		this.functionExcel = new ExcelReader("Configurations/Test_Cases/Functions.xlsx", sheetName);
		
	}
               
	/**
	 *
	 *@category: Main method of execution in framework, based on Flag name(Yes/No), the test case will be triggered in test suite excel sheet and get
	 *-----data in functions excel sheet 
	 * @return
	 */
               
	@Test
	public void getSuite(){
		ExcelReader suiteExcel = initializeSuiteExcel();
		Map<String, List<String>> testcaseID = Maps.newLinkedHashMap();
                               
		try{
			List<String> vals = Lists.newArrayList();
                                               
			for(int i = 1; i < suiteExcel.getNoOfRows(); i++){
				String testcaseName = suiteExcel.getDataByColumn(i, "Execute");
                                                               
				if(testcaseName.equalsIgnoreCase("Yes")){
                                                                               
					String caseID = suiteExcel.getDataByColumn(i, "Test_Case_ID");
 
					vals = suiteExcel.getColumnNames(i);
					System.out.println("**********vals :"+vals);
                                                                               
					testcaseID.put(caseID, vals);
                                                                               
				}
			}
                                               
			for(String name : testcaseID.keySet()){
				System.out.println("Testcase ID :"+name);
				Iterator bb = testcaseID.get(name).listIterator(3);
                                                               
				//Initialize Environment Excel Sheet
				int a=0;
				DataDrivenExcelSheet dd = null;
				if(a==0){
					dd = new DataDrivenExcelSheet();
					dd.setSheet(testcaseID.get(name).listIterator(3).next());
					
					a++;
				}
				while(bb.hasNext()){
                                                                               
					String parentName = bb.next().toString();
                                                               
					System.out.println("parentName :"+parentName);
                                                                               
					dataDriven(parentName,dd);
				}
			}
                                               
                                               
		}catch(Exception e){
			System.out.println("SheetName is not available in test suite");
			e.printStackTrace();
		}
		return;
	}
               
	/**
	 *
	 * @category: To initialize the Environment Suite excel sheet based on suite name by given in property file
	 */
               
	public void initializeEnvExcel(){
                               
		try{
			//Initialize Environment Excel Sheet
			envExcel = new ExcelReader("Configurations/Test_Setup/Test_Environment.xlsx", propValue.getProperty("environmentSheet"));
			System.out.println("Test Environment excel sheet is initialized successfully");
			
		}catch(Exception e){
                                               
			System.out.println("Environment sheet is not available in Configuration folder");
                                               
		}
	}
               
	/**
	 *
	 * @category: To initialize the TestSuite excel sheet based on suite name by given in property file
	 */
	public ExcelReader initializeSuiteExcel(){
		Properties propValue = initializePropertyFiles();
		try{
			//Initialize Environment Excel Sheet
			ExcelReader suiteExcel = new ExcelReader("Configurations/Test_Setup/Test_Suite.xlsx", propValue.getProperty("suiteName"));
			System.out.println("Test suite excel sheet is initialized successfully");
			System.out.println("*** suiteExcel ***:" +suiteExcel.getSheetName());
			return suiteExcel;
		}catch(Exception e){
                                               
			System.out.println("Suite excel sheet is not available in Configuration folder");
		}
		return envExcel;
	}
 
	/**
	 *
	 * @category: To initialize the configuration property files
	 */
               
	public Properties initializePropertyFiles(){
                               
		try{
			//Initialize Property File
			propValue = new Properties();
			fip = new FileInputStream(System.getProperty("user.dir")+"/src/com/framework/properties/config.properties");
			propValue.load(fip);
			System.out.println("Config Property file is loaded successfully");
		}catch(Exception e){
                                               
			System.out.println("Property file is not available");
		}
		return propValue;
	}
               
               
	BaseTestReport reports = new BaseTestReport();
               
	public HashMap<String, String> dataDriven(String parentSheetName,DataDrivenExcelSheet dd){
                               
		HashMap<String,String> map = new HashMap<>();
		this.functionExcel = dd.functionExcel;
		try{
                                               
                                               
			//Initialize Environment Excel Sheet
                                               
                                               
			System.out.println("Sheet Name Matched ");
                                                                               
			for(int i=1; i<functionExcel.getNoOfRows(); i++){
                                                                                               
				//Store the environment details in Hashmap
				String keywords = functionExcel.getDataByColumn(i, "Keywords");
                                                                                               
				if(parentSheetName.equalsIgnoreCase(keywords)){
//              	String testDesc = functionExcel.getDataByColumn(i, "Test_Description");
					String objects = functionExcel.getDataByColumn(i, "Objects_Reps");
					String events = functionExcel.getDataByColumn(i, "Events");
					String testData = functionExcel.getDataByColumn(i, "Test_Data");
					String reportdata = functionExcel.getDataByColumn(i, "Report_Status");
                                                                                               
					System.out.println("******************testData :"+testData);
                                                                                               
					reports.executeEvent(objects, events, testData, reportdata, driver);
                                                                                               
                                                                                               
                                                                                               
				}
			}
                                                                               
//       }
                                               
			return map;
                                               
                                               
		}catch(Exception e){
                                               
                                               
		}
		return null;
                               
	}
               
	public HashMap<String, String> getBrowserName(){
                               
		HashMap<String,String> map = new HashMap<>();
                               
		try{
                                               
			for(int i=1; i<envExcel.getNoOfRows(); i++){
				String statusCol = envExcel.getDataByColumn(i, "Status");
                                                               
				if(statusCol.equalsIgnoreCase("Yes")){
					//Store the environment details in Hashmap
					
					String browserName = envExcel.getDataByColumn(i, "BrowerName");
					String browserURL = envExcel.getDataByColumn(i, "URL");
					String testSuiteName = envExcel.getDataByColumn(i, "TestSuite_Name");
 
					map.put("browserName", browserName);
					map.put("browserURL", browserURL);
					map.put("testSuiteName", testSuiteName);
                                                                               
				}
                                                               
			}
			return map;
                                               
                                               
                                               
		}catch(Exception e){
                                               
                                               
                                               
		}
                               
                               
		return null;
	}
               
	/**
	 *
	 * @category: Initialize the browser by given name in Test Environment sheet
	 */
               
	@BeforeMethod
	public void initializeBrowser(){
                               
		String browserName = wrapper.getBrowserName().get("browserName");
		String browserURL = wrapper.getBrowserName().get("browserURL");
		
		try{
                                	
			System.out.println("browserName :"+browserName);
			System.out.println("browserURL :"+browserURL);
			
			if(browserName.equalsIgnoreCase("Chrome")){             
                                               
				System.setProperty("webdriver.chrome.driver", "BrowserDrivers/chromedriver.exe");
				driver = new ChromeDriver();
				ExtentTestManager.getTest().log(LogStatus.INFO, ""+browserName+" browser is initialized successfully");
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.get(browserURL);
				ExtentTestManager.getTest().log(LogStatus.INFO, ""+browserURL+" browser URL is loaded successfully");
                               
			}else if(browserName.equalsIgnoreCase("IE")){
                                               
				System.setProperty("webdriver.ie.driver", "BrowserDrivers/IEDriverServer.exe");
				driver = new InternetExplorerDriver();
				ExtentTestManager.getTest().log(LogStatus.INFO, ""+browserName+" browser is initialized successfully");
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.get(browserURL);
				ExtentTestManager.getTest().log(LogStatus.INFO, ""+browserURL+" browser URL is loaded successfully");
                                               
			}else if(browserName.equalsIgnoreCase("Firefox")){
                                               
				driver = new FirefoxDriver();
				ExtentTestManager.getTest().log(LogStatus.INFO, ""+browserName+" browser is initialized successfully");
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.get(browserURL);
				ExtentTestManager.getTest().log(LogStatus.INFO, ""+browserURL+" browser URL is loaded successfully");
                                               
			}
		}catch(Exception e){
                                               
			e.printStackTrace();
                                               
		}
                               
	}
               
	@BeforeSuite
	public void initializeFiles(){
                               
		try{
                               
			//Initialize property files
                               
			wrapper.initializePropertyFiles();
                                                               
			//Initialize and get Values from Excel suite
			wrapper.initializeSuiteExcel();
			wrapper.initializeEnvExcel();
                               
			//Kill browser driver exe's
			wrapper.killBrowserDrivers();
                               
			//Kill Browser exe's
			wrapper.killBrowsers();
                               
		}catch(Exception e){
                               
			e.printStackTrace();
                               
		}
                               
                               
	}
               
	public void killBrowserDrivers(){
                               
		try{
			String browserName = wrapper.getBrowserName().get("browserName");
                                               
			if(browserName.equalsIgnoreCase("Chrome")){
				Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe");
				System.out.println("Chrome driver .exe process is deleted successfully in task manager");
			}else if(browserName.equalsIgnoreCase("Firefox")){
				Runtime.getRuntime().exec("taskkill /F /IM geckodriver.exe");
				System.out.println("Gecko driver .exe process is deleted successfully in task manager");
			}else if(browserName.equalsIgnoreCase("IE")){
				Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
				System.out.println("IEDriverServer .exe process is deleted successfully in task manager");
			}
                                                               
		}catch(Exception e){
			e.printStackTrace();
		}
                               
                               
                               
	}
               
	public void killBrowsers(){
                               
		try{
                                               
			String killBrowserCondition = propValue.getProperty("killBrowser");
                                               
			String browserName = wrapper.getBrowserName().get("browserName");
                                               
			if(killBrowserCondition.equalsIgnoreCase("True")){
				System.out.println("Kill browser condition value is True");
				System.out.println("Delete the browser driver.exe before start the execution");
                                                               
				if(browserName.equalsIgnoreCase("Chrome")){
					Runtime.getRuntime().exec("taskkill /F /IM chrome.exe");
					System.out.println("Chrome browser is killed successfully");
				}else if(browserName.equalsIgnoreCase("Firefox")){
					System.out.println("Firefox browser is killed successfully");
					Runtime.getRuntime().exec("taskkill /F /IM firefox.exe");
				}else if(browserName.equalsIgnoreCase("IE")){
					Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
					System.out.println("IE browser is killed successfully");
                                                                }
                                               
			}else{
                                                               
				System.out.println("Kill browser condition value is False");
                                                               
			}
                                                               
		}catch(Exception e){
			e.printStackTrace();
		}
	}
               
	@AfterTest
	public void quitDriver(){
                               
		driver.close();
                               
	}
               
	@AfterSuite
	public void afterSuiteMethod(){
                               
                               
		killBrowserDrivers();
                               
	}
               
               
	public WebDriver getDriver(){
                               
		return driver;
                               
	}
               
               
	public static void main(String[] args) {
                               
		try{
                               
                               
			//Initialize property files
			wrapper.initializePropertyFiles();
                               
			//Initialize and get Values from Excel suite
			wrapper.initializeSuiteExcel();
			wrapper.initializeEnvExcel();
			wrapper.killBrowserDrivers();
			wrapper.killBrowsers();
			wrapper.initializeBrowser();
			wrapper.getSuite();
                               
//          wrapper.initializeReports();
                               
		}catch(Exception e){
			e.printStackTrace();
		}
	}
}