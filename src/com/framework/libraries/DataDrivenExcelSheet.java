package com.framework.libraries;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class DataDrivenExcelSheet {
	
	public ExcelReader functionExcel;
	public WebDriver driver = null;
	
	
	
	public void setSheet(String sheetName) throws FileNotFoundException, IOException{
		this.functionExcel = new ExcelReader("Configurations/Test_Cases/Functions.xlsx", sheetName);

	}
	
	
	/**
	 * 
	 * @category: This method is to get the data from excel sheet to execute the test case execution
	 * @param parentSheetName
	 * @param dd
	 * @return
	 */
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
						String testDesc = functionExcel.getDataByColumn(i, "Test_Description");
						String objects = functionExcel.getDataByColumn(i, "Objects_Reps");
						String events = functionExcel.getDataByColumn(i, "Events");
						String testData = functionExcel.getDataByColumn(i, "Test_Data");
						
//						executeEvent(objects, events, testData);
						
						map.put("keywords", keywords);
						map.put("testDesc", testDesc);
						map.put("objects", objects);
						map.put("events", events);
						map.put("testData", testData);
						
						System.out.println(map);
						}
					}
					
//			}
			
			return map;
			
			
		}catch(Exception e){
			
			
		}
		return null;
		
		
	}
	
	/**
	 * 
	 * @category: This method is to execute the test case steps by given event names(Keyword) and split the locator type and locator by '#'
	 * @param locator
	 * @param eventName
	 * @param testdata
	 * @return
	 */
	public String executeEvent(String locator,String eventName,String testdata){
		
		try{
			
			String locatorTypeValueFromObj = locator.split("#")[0];
			String locatorValueFromObj = locator.split("#")[1];
			
			WebElement element = byLocatorType(locatorTypeValueFromObj, locatorValueFromObj);
			
			switch(eventName.toLowerCase()){
			
			case "Enter":
				element.clear();
				element.sendKeys(testdata);
				return "Typed the value";
				
			case "Click":
				element.click();
				return "Clicked the value";
				
				
			}
			
			
			}catch(Exception e){
				
				e.printStackTrace();
				
			}
			
			return null;
		
	}
	
	/**
	 * 
	 * @category: This method is to find the locator type and locator and used in 'Execute Event' method
	 * @param locatorType
	 * @param locator
	 * @return
	 */
	
	public WebElement byLocatorType(String locatorType, String locator){
		
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
	
}
