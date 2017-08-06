package com.framework.libraries;

import java.io.File;

import com.relevantcodes.extentreports.ExtentReports;

public class ExtentManager {
	private static ExtentReports instance;
	
	public static synchronized ExtentReports getInstance() {
		if (instance == null) {
			
			String reportPath = "Reports/Extent_Reports/TestReport.html";
			
			
			System.out.println(System.getProperty("user.dir"));
			instance = new ExtentReports(reportPath);
			
			instance.addSystemInfo("HostName", "Praveen")
			.addSystemInfo("Environment", "Report Creation")
			.addSystemInfo("UserName", "Praveen Lakshmanan");
			instance.loadConfig(new File("SupportedFiles/extent-config.xml"));
			
		}
		
		return instance;
	}
}