package com.framework.libraries;
 
import java.io.File;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
 
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.testng.TestListenerAdapter;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
 
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
 
	public class BaseTestReport extends TestListenerAdapter {
 
		public ExtentTest testReporter;
               
		@Parameters({"testcaseID", "testcaseDesc"})
		@BeforeMethod
		public void beforeMethod(Method caller, String testcaseID, String testcaseDesc){
                               
//        ExtentTestManager.startTest(caller.getName(), "This is a sample test.");
			ExtentTestManager.startTest(testcaseID, testcaseDesc);
                               
                }
               
		public String executeEvent(String locator, String eventName, String testdata, String reportName, WebDriver driver){
                               
                               
			try{
				String locatorTypeValueFromObj = locator.split("#")[0];
				String locatorValueFromObj = locator.split("#")[1];
                               
				WebElement element = byLocatorType(locatorTypeValueFromObj, locatorValueFromObj, driver);
                               
				switch(eventName.toLowerCase()){
                               
				case "enter":
                                               
					try{
						element.clear();
						element.sendKeys(testdata);
						ExtentTestManager.getTest().log(LogStatus.PASS, "Value '"+testdata+"' are entered successfully on '"+reportName+"'");
						return "Typed the value";
						
					}catch(Exception e){
						
						ExtentTestManager.getTest().log(LogStatus.FAIL, "Exception occurred, while enter the '"+testdata+"' value on element");
						ExtentTestManager.getTest().log(LogStatus.FAIL, ExtentTestManager.getTest().addScreenCapture(reportFailScreenshot(driver)));
						
                                               
					}
                                               
					break;
                                               
				case "click":
                                               
					try{
						element.click();
						ExtentTestManager.getTest().log(LogStatus.PASS, "Element '"+reportName+"' is clicked successfully");
						return "Clicked the value";
                                               
					}catch(Exception e){
                                                               
						ExtentTestManager.getTest().log(LogStatus.FAIL, "Exception occurred, while clicking on element");
						ExtentTestManager.getTest().log(LogStatus.FAIL, ExtentTestManager.getTest().addScreenCapture(reportFailScreenshot(driver)));
                                                               
					}             
                                               
					break;
                                               
				}
                               
                               
			}catch(Exception e){
                                               
				e.printStackTrace();
                                               
			}
                               
                               
                               
			return null;
                               
		}
               
               
		public WebElement byLocatorType(String locatorType, String locator, WebDriver driver){
                               
			WebElement element = null;
                               
			try{
				switch (locatorType.toLowerCase()) {
                                               
				case "id":
					element = driver.findElement(By.id(locator));
					break;
                                                               
				case "xpath":
					element = driver.findElement(By.xpath(locator));
					break;
 
				case "css":
					element = driver.findElement(By.cssSelector(locator));
					break;
                                                               
				case "className":
					element = driver.findElement(By.className(locator));
					break;
                                                               
				case "partialLinkText":
					element = driver.findElement(By.partialLinkText(locator));
					break;
                                                               
				case "tagName":
					element = driver.findElement(By.tagName(locator));
					break;
                                               
				case "linkText":
					element = driver.findElement(By.linkText(locator));
					break;
                                                               
				default:
					throw new IllegalArgumentException("Given Locator Typs is not available in findBy mechanism in webdriver");
                                                                
				}
                                               
			}catch(Exception e){
				e.printStackTrace();
			}
			return element;
		}
               
               
		public String reportFailScreenshot(WebDriver driver){
                               
                               
			try{
                                               
				String screenshotPath = "Reports/Extent_Reports/Screenshots/";
				if(!(new File(screenshotPath).exists())){
					System.out.println("Screenshot folder is not available, it's going to create the directory");
					new File(screenshotPath).mkdir();
					System.out.println("Screenshot folder is created successfully");
				}
                               
				String screenshotName = new SimpleDateFormat("dd_MM_yyyy_hh_mm_ss_a").format(new Date());
				File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(screenshotFile, new File(screenshotPath+screenshotName+".png"));
//             	System.out.println("screenshotName :"+screenshotName);
				File fileName = new File(screenshotPath+screenshotName+".png");
				String absolutePath = fileName.getAbsolutePath();
				System.out.println("absolutePath :"+absolutePath);
				return absolutePath;
                               
                               
			}catch(Exception e){
				e.printStackTrace();
			}
			return null;
                               
		}
 
		@AfterSuite
		public void afterSuite(){
			
			ExtentTestManager.endTest();
			ExtentManager.getInstance().flush();
                               
		}
               
}
