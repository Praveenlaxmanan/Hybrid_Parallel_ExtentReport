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


Use: 
ReadExcel:

package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import Fillo.Connection;
import Fillo.Fillo;
import Fillo.Recordset;


public class ReadExcel {

	public static String filename;
	public  String path;
	public  FileInputStream fis = null;
	public  FileOutputStream fileOut =null;
	private HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null;
	private HSSFRow row   =null;
	private HSSFCell cell = null;

	public ReadExcel(String path){
		this.path=path;

		try{
			fis = new FileInputStream(path);
			//fileOut = new FileOutputStream(path);
			workbook = new HSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		}catch(Exception e){
			e.printStackTrace();
		}
	}

	public int getRowCount(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1)
			return 0;
		else{
			sheet = workbook.getSheetAt(index);
			int number=sheet.getLastRowNum();
			return number;

		}
	}

	public int getColumnCount(String sheetName){

		int index = workbook.getSheetIndex(sheetName);
		if(index==-1)
			return 0;
		else{
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(0);
		}

		if(row==null)
			return -1;

		return row.getLastCellNum();
	}

	public String getCellData(String sheetName,String colName,int rowNum){
		try{
			if(rowNum <=0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			if(index==-1)
				return "";

			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(0);

			int col_Num=-1;
			for(int i=0;i<row.getLastCellNum();i++){
				if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num=i;
			}
			if(col_Num==-1)
				return null;

			row = sheet.getRow(rowNum);
			if(row==null)
				return "";
			cell = row.getCell(col_Num);

			if(cell==null)
				return "";

			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();

			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC )
				return String.valueOf(cell.getNumericCellValue()).trim();

			else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue()).trim();

		}
		catch(Exception e){
			e.printStackTrace();
			// return "row "+rowNum+" or column "+colName +" does not exist in xls";
			return "";
		}

	}

	public int getNumericCellData(String sheetName,String colName,int rowNum){
		int value = -1;
		try{
			if(rowNum <=0)
				return -1;

			int index = workbook.getSheetIndex(sheetName);
			if(index==-1)
				return -1;

			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(0);

			int col_Num=-1;
			for(int i=0;i<row.getLastCellNum();i++){
				if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num=i;
			}
			if(col_Num==-1)
				return -1;

			row = sheet.getRow(rowNum);
			if(row==null)
				return -1;
			cell = row.getCell(col_Num);

			if(cell==null)
				return -1;

			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC )
				value = (int) (cell.getNumericCellValue());

			else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return -1;
		}
		catch(Exception e){
			e.printStackTrace();
			return -1;
		}
		return value;

	}

	public String getCellData(String sheetName,int colNum,int rowNum){
		try{
			if(rowNum <=0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			if(index==-1)
				return "";

			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(0);

			row = sheet.getRow(rowNum);
			if(row==null)
				return "";
			cell = row.getCell(colNum);

			if(cell==null)
				return "";

			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();

			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC )
				return String.valueOf(cell.getNumericCellValue());

			else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());

		}
		catch(Exception e){
			e.printStackTrace();
			return "row "+rowNum+" or column "+colNum +" does not exist  in xls";
		}

	}

	public String RetrieveAutomationKeyFromExcel(String wsName, String colName, String rowName) throws Exception{

		try{

			int index = workbook.getSheetIndex(wsName);
			if(index==-1)
				return "";

			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(1);

			HSSFRow headRow=sheet.getRow(1);

			int rowCount = getRowCount(wsName);
			int colCount = getColumnCount(wsName);
			int colNumber=-1;	
			int rowNumber=-1;	

			for(int i=0; i<colCount; i++){
				if(headRow.getCell(i).getStringCellValue().equals(colName.trim())){
					colNumber=i;
					break;
				}					
			}
			if(colNumber==-1){
				throw new RuntimeException("colNumber  -1");				
			}

			for(int j=0; j<rowCount; j++){
				HSSFRow Suitecol = sheet.getRow(j);				
				if(Suitecol.getCell(1).getStringCellValue().equals(rowName.trim())){
					rowNumber=j;
					break;
				}					
			}

			if(rowNumber==-1){
				return "";				
			}

			row = sheet.getRow(rowNumber);
			cell = row.getCell(colNumber);
			if(cell==null)
				return "";

			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue().trim();

			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC )
				return String.valueOf(cell.getNumericCellValue()).trim();

			else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue()).trim();

		}catch(Exception e){
			e.printStackTrace();
			throw e;
		}
	}

	public boolean CheckTestCaseNamedSheetExist(String filelocation, String TESTCASE_NAME) throws Exception{
		FileInputStream ipstr = null;
		HSSFWorkbook wb = null;
		boolean sheetExist=false;
		try{
			ipstr = new FileInputStream(filelocation);
			wb = new HSSFWorkbook(ipstr);
			ipstr.close();
			for(int k=0; k<wb.getNumberOfSheets(); k++){
				if(wb.getSheetName(k).equalsIgnoreCase(TESTCASE_NAME)){
					sheetExist=true;
				}
			}
		}catch(Exception e){
			throw e;
		}finally{
			ipstr = null;
			if(wb!=null){
				wb.close();
				wb=null;
			}
		}
		return sheetExist;
	}

		public static String RetrieveValueFromTestDataBasedOnQuery(String FileLocation,String FileName,String Select_Column_Name,String sqlQuery) throws Exception{

		Fillo fillo=null;
		Connection filloConnection = null;
		Recordset recordset=null;
		String filePath = null;
		String record=null;
		
		try {
			filePath = System.getProperty("user.dir")+FileLocation+FileName;
			fillo=new Fillo();
			filloConnection=fillo.getConnection(filePath);
			recordset=filloConnection.executeQuery(sqlQuery);
			while(recordset.next()){
			record=recordset.getField(Select_Column_Name);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		finally{
		fillo = null;
		filloConnection.close();
		}
		
		return record;
		
	}
}


public static void readVBSFile(){
		
		try{
		
		String filePath = "chrome/read.vbs";
		File file = new File(filePath);
		String s = IOUtils.toString(file.toURI(), "UTF-8");
//		System.out.println(s);
//		Pattern patter = Pattern.compile("(DynamicDO.Add) (RunTimeMasterNumber_TS01), (2356925684)");
		Pattern patter = Pattern.compile("DynamicDO.Add \"RunTimeMasterNumber_TS07\", \"(.*)\"");
		Matcher matcher = patter.matcher(s);
		while(matcher.find()){
			
			System.out.println(matcher.group(1));
			}
		}catch(Exception e){
			
			e.printStackTrace();
		}
	}
