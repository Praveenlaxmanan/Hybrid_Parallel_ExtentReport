package com.framework.libraries;
 
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
 
import org.testng.TestNG;
import org.testng.annotations.Test;
 
	public class TestNGRunner {
 
		@Test
		public void testReport(){
                               
                               
			List<Class> testCases = new ArrayList<Class>();
			testCases.add(Wrappers.class);
			
			ExecutorService execService = Executors.newFixedThreadPool(2);
                               
			for(final Class testFile : testCases){
                                               
				execService.submit(new Runnable() {
					
					@Override
					public void run() {
						
						System.out.println("Running test file :"+testFile.getName());
						testRunnerTestNG(testFile);
                                                                }
				});
                                               
			}
			execService.shutdown();
			try{
				
				execService.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);
                                               
			}catch(Exception e){
                                               
				e.printStackTrace();
			}
                               
			System.out.println("Completing the execution service");
                              
		}
               
		public static void testRunnerTestNG(Class args) {
                               
			TestNG testNG = new TestNG();
			testNG.setTestClasses(new Class[] { args });
			testNG.run();
                               
		}
}
