package com.framework.testNG;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.framework.libraries.ExcelReader;

public class CreateXMLs {
	public ExcelReader getSuiteExcel;
	
	private static Logger logInfo = Logger.getLogger(CreateXMLs.class.getName());
	private static Properties parallelConfig;
	
	static int testcaseXMLSplitCount = 0;
	
	
	public void createTestcaseXML(String testcaseID, String testcaseDesc, String moduleName, String testAuthor){
		
		DocumentBuilderFactory docBulFactory = null;
		DocumentBuilder docBuilder = null;
		Document document = null;
		Element rootElement = null;
		
		
		try{
			//Instantiate the DOM to create XML
			docBulFactory = DocumentBuilderFactory.newInstance();
			docBuilder = docBulFactory.newDocumentBuilder();
			document = docBuilder.newDocument();
			
			
			rootElement = document.createElement("suite");
			rootElement.setAttribute("name", "Testcases");
			rootElement.setAttribute("parallel", "tests");
			rootElement.setAttribute("thread-count", "3");
			
			Element parameterName = document.createElement("parameter");
			parameterName.setAttribute("name", "suiteType");
			parameterName.setAttribute("value", "Parallel");
			rootElement.appendChild(parameterName);
			document.appendChild(rootElement);
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
			
			
		try{
		
			//Append test case details in TestNG XML files
			//Create Headers in Parallel TestNG file for Parallel Suite
			
			Element test = document.createElement("test");
			test.setAttribute("name", testcaseID);
			rootElement.appendChild(test);
			
			Element parameter = document.createElement("parameter");
			parameter.setAttribute("name", "testcaseID");
			parameter.setAttribute("value", testcaseID);
			test.appendChild(parameter);
			
			Element parameter2 = document.createElement("parameter");
			parameter2.setAttribute("name", "testcaseDesc");
			parameter2.setAttribute("value", testcaseDesc);
			test.appendChild(parameter2);
			
			Element parameter3 = document.createElement("parameter");
			parameter3.setAttribute("name", "moduleName");
			parameter3.setAttribute("value", moduleName);
			test.appendChild(parameter3);
			
			Element parameter4 = document.createElement("parameter");
			parameter4.setAttribute("name", "testAuthor");
			parameter4.setAttribute("value", testAuthor);
			test.appendChild(parameter4);
			
			Element classes = document.createElement("classes");
			test.appendChild(classes);
			
			Element singleClass = document.createElement("class");
			singleClass.setAttribute("name", "com.framework.libraries.Wrappers");
			classes.appendChild(singleClass);
			
			Element method = document.createElement("methods");
			singleClass.appendChild(method);
			
			Element include = document.createElement("include");
			include.setAttribute("name", "getSuite");
			method.appendChild(include);
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
			
		try{
			
			//Create the TestNG XML file with DOC object
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty(OutputKeys.DOCTYPE_SYSTEM, "http://testng.org/testng-1.0.dtd");
			DOMSource domSource = new DOMSource(document);
			
			String testNGXMLFilePath = "TestNG_XMLs/testCreate.xml";
			
			StreamResult result = new StreamResult(new File(testNGXMLFilePath));
			
			transformer.transform(domSource, result);
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
			
		
		
		
		
	}
	
	public void initializeSuiteExcel(){
		
		try{
		//Initialize Environment Excel Sheet
			getSuiteExcel = new ExcelReader("Configurations/Test_Setup/Test_Suite.xlsx", parallelConfig.getProperty("suiteName"));
		System.out.println("*** suiteExcel ***:" +getSuiteExcel.getSheetName());
		}catch(Exception e){
			
			System.out.println("Suite excel sheet is not available in Configuration folder");
		}
	}
	
	
	public void createTestNGXmlFile(){
		
		
		//Instantiate the DOM to create XML
		DocumentBuilderFactory docBulFactory = null;
		DocumentBuilder docBuilder = null;
		Document document = null;
		Element rootElement = null;
		
		
		try{
			docBulFactory = DocumentBuilderFactory.newInstance();
			docBuilder = docBulFactory.newDocumentBuilder();
			document = docBuilder.newDocument();
			
			
			rootElement = document.createElement("suite");
			rootElement.setAttribute("name", "Testcases");
			rootElement.setAttribute("parallel", "tests");
			rootElement.setAttribute("thread-count", "3");
			
			Element parameterName = document.createElement("parameter");
			parameterName.setAttribute("name", "suiteType");
			parameterName.setAttribute("value", "Parallel");
			rootElement.appendChild(parameterName);
			document.appendChild(rootElement);
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
		
		try{
			
			for(int i = 1; i < getSuiteExcel.getNoOfRows(); i++){
				
				String executeFlag = getSuiteExcel.getDataByColumn(i, "Execute");
				
				if(executeFlag.equalsIgnoreCase("Yes")){
					
					String testcaseID = getSuiteExcel.getDataByColumn(i, "Test_Case_ID");
					String executionType = getSuiteExcel.getDataByColumn(i, "Execution_Type");
					String testcaseDesc = getSuiteExcel.getDataByColumn(i, "Testcase_Description");
					String moduleName = getSuiteExcel.getDataByColumn(i, "Component_SheetName");
					String testAuthor = getSuiteExcel.getDataByColumn(i, "TestAuthor");
					
					System.out.println("caseID :"+testcaseID);
					System.out.println("executionType :"+executionType);
					System.out.println("testcaseDesc :"+testcaseDesc);
					System.out.println("moduleName :"+moduleName);
					System.out.println("moduleName :"+testAuthor);
					
					
					//Append test case details in TestNG XML files
					//Create Headers in Parallel TestNG file for Parallel Suite
					
					Element test = document.createElement("test");
					test.setAttribute("name", testcaseID);
					rootElement.appendChild(test);
					
					Element parameter = document.createElement("parameter");
					parameter.setAttribute("name", "testcaseID");
					parameter.setAttribute("value", testcaseID);
					test.appendChild(parameter);
					
					Element parameter2 = document.createElement("parameter");
					parameter2.setAttribute("name", "testcaseDesc");
					parameter2.setAttribute("value", testcaseDesc);
					test.appendChild(parameter2);
					
					Element parameter3 = document.createElement("parameter");
					parameter3.setAttribute("name", "moduleName");
					parameter3.setAttribute("value", moduleName);
					test.appendChild(parameter3);
					
					Element parameter4 = document.createElement("parameter");
					parameter4.setAttribute("name", "testAuthor");
					parameter4.setAttribute("value", testAuthor);
					test.appendChild(parameter4);
					
					Element classes = document.createElement("classes");
					test.appendChild(classes);
					
					Element singleClass = document.createElement("class");
					singleClass.setAttribute("name", "com.framework.libraries.Wrappers");
					classes.appendChild(singleClass);
					
					Element method = document.createElement("methods");
					singleClass.appendChild(method);
					
					Element include = document.createElement("include");
					include.setAttribute("name", "getSuite");
					method.appendChild(include);
					
				}
				
			}
			
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
		
		
		try{
			
			//Create the TestNG XML file with DOC object
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty(OutputKeys.DOCTYPE_SYSTEM, "http://testng.org/testng-1.0.dtd");
			DOMSource domSource = new DOMSource(document);
			
			String testNGXMLFilePath = "TestNG_XMLs/testCreate.xml";
			
			StreamResult result = new StreamResult(new File(testNGXMLFilePath));
			
			transformer.transform(domSource, result);
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
		
		
	}
	
	public static void main(String[] args) {
		
		try{
			
			//Initialized parallel configuration property file
			parallelConfig = new Properties();
			parallelConfig.load(new FileInputStream(new File("./src/com/framework/properties/parallelConfig.properties")));
			
			//Initialized Log4j config file
			String log4jConfigPath = parallelConfig.getProperty("Log4jConfigFile");
			PropertyConfigurator.configure(log4jConfigPath);
			logInfo.info("********* Convert Test Cases to TestNG XML Files ***********\n");
			
			//Split the test cases in XML by given in parallel configuration property file (Not Implemented)
			String testcaseSplitCount = parallelConfig.getProperty("testCaseSplitCount");
			testcaseXMLSplitCount = Integer.parseInt(testcaseSplitCount);
			logInfo.info("Splitted test case count :"+testcaseXMLSplitCount);
			
		}catch(Exception e){
			
			e.printStackTrace();
			
		}
		
		CreateXMLs test = new CreateXMLs();
		test.initializeSuiteExcel();
		test.createTestNGXmlFile();
		
		
	}
	

}
