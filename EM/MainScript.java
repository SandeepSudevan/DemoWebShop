package EM;

import java.awt.AWTException;

/*Description: Access TestCase and TestSet Sheets.
			   Read required TestCase row by row and compare the elements & does action on it.*/


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import javax.swing.JComboBox;
import javax.swing.JTable;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.UnexpectedAlertBehaviour;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.chromium.ChromiumDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;



public class MainScript{
	 //********************Variable declaration************************************
	int ExecuteTotalRows,row = 0,col=0;
	static int ExecuteTotalCol,TestCaseRow,TestScenarioTotalRow,TestScenarioRow,Count,randnum,winnumber,p,n,celltype,p2,p3,p4,p5,modcol=0;
	static String End_Time,Flag,Browserdec,Browsersec;
	static String Start_Time,Password,Username,Facility;
	static String Execute,FolderName,cellvalue,value,Parentframename=null,iframename=null,framename=null,frameinsideframesname=null,parenthandle,date1,TestScenario,Responsibility=null,DataSelection,Module,title,ModFunc,time1,result,Object1,ProName1,ProValue1,AddProname,AddProValue,ID,comments,Parameter1,Input,filePath;
    public static WebDriver driver;
	static Double cellvalueformula;
	static long Elapsed_Time_in_seconds;
	static HSSFSheet Step_Report,Exectimesh,CheckPointsh;
	static HSSFWorkbook workbook,TestCaseWorkbookIN;
	static FormulaEvaluator evaluator;
	static FileInputStream TestCaseExcelIN;
	static WebElement ele;
	@SuppressWarnings({ "resource", "unused","rawtypes"})
	
	public static void launch(JComboBox comboBox,JComboBox comboBox_1,JComboBox comboBox_2,JTextField textField,JTextField textField_1,java.awt.Component comp,JTable table_1) throws BiffException, IOException, InterruptedException ,NoAlertPresentException, WriteException
	{
		try
		{
		driver.manage().deleteAllCookies();
		}
		catch(Exception e){}
		String FolderName=comboBox.getSelectedItem().toString();
		String SiteName=comboBox_1.getSelectedItem().toString();
		String Responsibility=comboBox_2.getSelectedItem().toString();
		int TestScenarioRow=Integer.parseInt(textField.getText());
		int TestScenarioEndRow=Integer.parseInt(textField_1.getText());
		JavascriptExecutor jsExecutor = (JavascriptExecutor)driver;
		HSSFWorkbook workbook = null;
		java.util.Date date = new java.util.Date();
		String filePath ="C:/Users/mchalla3/workspace latest/Enterprise Managment/Results/ReportExcel.xls";
		FileInputStream inputFile = new FileInputStream(filePath);
		workbook = new HSSFWorkbook(inputFile);
		HSSFSheet[] Resultsheets=ResultReporting.ReportExcel(filePath,workbook);
		Step_Report=Resultsheets[0];
		CheckPointsh=Resultsheets[1];
		Exectimesh=Resultsheets[2];
		String parentWindow = null;
		//System.setProperty("webdriver.ie.driver", "C:/Users/mchalla3/eclipse/IEDriverServer.exe");
//		System.setProperty("webdriver.chrome.driver", "C:/Users/mchalla3/workspace latest/chromedriver-win32/chromedriver.exe");
		System.setProperty("webdriver.chrome.driver", "C:/Users/mchalla3/workspace latest/chromedriver-win64/chromedriver.exe");
//		System.setProperty("webdriver.bat.driver", "C:/eclipse-jee-neon-3-win32-x86_64/eclipse/EM12XSEL.bat");
//     System.setProperty("webdriver.edge.driver", "C:/eclipse-jee-neon-3-win32-x86_64/eclipse/msedgedriver.exe");
		
//		InternetExplorerOptions ieOptions = new InternetExplorerOptions();
//		ieOptions.attachToEdgeChrome();
//		ieOptions.withEdgeExecutablePath("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"); // Edge Browser directory path
//		ieOptions.setCapability("ignoreProtectedModeSettings", true);
//		ieOptions.setUnhandledPromptBehaviour(UnexpectedAlertBehaviour.IGNORE);
//		
//		driver = new InternetExplorerDriver(ieOptions);

//		driver.navigate().to("http://130.78.62.220:9999/HIS/eSM/jsp/login.jsp");//em12xsut
//driver.findElement(By.name("name")).sendKeys("CSCCA");
//driver.findElement(By.name("password")).sendKeys("CSCCA");
//WebElement loginbtn = (new WebDriverWait(driver,Duration.ofSeconds(5)))
//		  .until(ExpectedConditions.presenceOfElementLocated(By.id("loginID")));
//loginbtn.click();
//driver.findElement(By.className("dhx_combo_input")).sendKeys(Responsibility);
//driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.DOWN);
//driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.ENTER);
//driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.TAB);
//driver.findElement(By.id("loginID")).click();

//		ChromeOptions opt = new ChromeOptions();
//		opt.setExperimentalOption("excludeSwitches",Arrays.asList("enable-automation"));
		
		ChromeOptions opt = new ChromeOptions();
		opt.addArguments("--remote-allow-origins=*");
//		driver = new InternetExplorerDriver();
//		driver = new EdgeDriver();
		driver = new ChromeDriver(opt);
		
//driver.navigate().to("http://130.78.62.220:9999/HIS/eCommon/jsp/eHIS.jsp");
//		driver.navigate().to("http://130.78.63.1:8989/HIS/eSM/jsp/login.jsp");
//		driver.navigate().to("http://130.78.63.1:8989/HIS/eSM/jsp/login.jsp?computerName=%22DUC-MVa1lIs6W0u%22&ipAddress=%22192.168.0.203%22");
	
		
		WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(10));
 	   	WebDriverWait waitele = new WebDriverWait(driver,Duration.ofSeconds(5));
 	   	WebDriverWait waitframe=new WebDriverWait(driver,Duration.ofSeconds(30));
 		WebDriverWait waitwin=new WebDriverWait(driver,Duration.ofSeconds(2));
 	   	DataFormatter formatter = new DataFormatter();
 	 try
		{
 		 	String testcaseExcel="C:/Users/mchalla3/workspace latest/Enterprise Managment/InputFiles/EM_Automation_Test Case.xls";
 		 	TestCaseExcelIN= new FileInputStream(new File(testcaseExcel));
 		 	TestCaseWorkbookIN= new HSSFWorkbook(TestCaseExcelIN);
 		 	String ScenarioExcel = "C:/Users/mchalla3/workspace latest/Enterprise Managment/InputFiles/EM_Automation_Test_Set.xls";
			FileInputStream TestScenarioExcelIN= new FileInputStream(new File(ScenarioExcel));
            HSSFWorkbook TestScenarioWorkbookIN= new HSSFWorkbook(TestScenarioExcelIN);
            HSSFSheet TestScenarioSheetIN=TestScenarioWorkbookIN.getSheet(FolderName);
	 		TestScenarioTotalRow=TestScenarioSheetIN.getLastRowNum();
//	 		loginPage(Responsibility,parentWindow,wait,"cscca","cscca","Siriraj Hospital");
	 		//loginPage(Responsibility,parentWindow,wait,"CSCBL","CSCBL","Siriraj Hospital");
	 		loginPage(Responsibility,parentWindow,wait,"noppadolth","sihmis","Siriraj Hospital");	 		
	 		for(int k=TestScenarioRow;k<=TestScenarioEndRow;k++)		
	 		{
	 			try{
	 				if(!(TestScenarioSheetIN.getRow(TestScenarioRow).getCell(0).getStringCellValue()==null) && TestScenarioSheetIN.getRow(TestScenarioRow).getCell(0).getStringCellValue().length() != 0)
	 				{
	 					TestScenario = TestScenarioSheetIN.getRow(TestScenarioRow).getCell(0).getStringCellValue();
	 					System.out.println("***********************TestScenarioRow***********************"+TestScenarioRow);
	 					System.out.println("***********************TestScenario***********************"+TestScenario);
	 					
	 				}
	 			}
	 			catch(Exception e){}
	 			try{
	 				if(!(TestScenarioSheetIN.getRow(TestScenarioRow).getCell(1).getStringCellValue()==null) && TestScenarioSheetIN.getRow(TestScenarioRow).getCell(0).getStringCellValue().length() != 0)
	 				{
	 					Responsibility=TestScenarioSheetIN.getRow(TestScenarioRow).getCell(1).getStringCellValue();
	 				}
	 			}
	 			catch(Exception e){}
	 			try{
	 				if(!(TestScenarioSheetIN.getRow(TestScenarioRow).getCell(2).getStringCellValue()==null) && TestScenarioSheetIN.getRow(TestScenarioRow).getCell(0).getStringCellValue().length() != 0)
	 				{
	 					DataSelection=TestScenarioSheetIN.getRow(TestScenarioRow).getCell(2).getStringCellValue();
	 				}
	 			}
	 			catch(Exception e){}
	 			try{
	 				Username = TestScenarioSheetIN.getRow(TestScenarioRow).getCell(3).getStringCellValue();
	 				if(Username==null || Username.equals(""))
	 				{
	 					Username ="noppadolth";
	 				}
	 				
	 			}
	 			catch(Exception e){Username ="noppadolth";}
	 			try{
	 				Password = TestScenarioSheetIN.getRow(TestScenarioRow).getCell(4).getStringCellValue();
	 				if(Password==null || Password.equals(""))
	 				{
	 					Password ="sihmis";
	 				}
	 				
	 			}
	 			catch(Exception e){Password ="sihmis";}
	 			try{
	 				Facility = TestScenarioSheetIN.getRow(TestScenarioRow).getCell(5).getStringCellValue();
	 				
	 				if(Facility==null || Facility.equals("") || Facility.equalsIgnoreCase("common"))
	 				{
	 					Facility ="Women and Child Care hospital";
	 					//Facility ="SELAYANG HOSPITALS";
	 				}
	 			
	 			else{Facility =Facility;}
	 				}
	 				catch(Exception e){};
	 			HSSFSheet TestCaseSheetIN=TestCaseWorkbookIN.getSheet(TestScenario);
		        HSSFSheet TestCaseSheetMenu=TestCaseWorkbookIN.getSheet("Menu");
		        int TestcaseTotalRow = TestCaseSheetIN.getPhysicalNumberOfRows();
		 		Start_Time=Starttime();
		 		
		 		switch(TestScenario)
		 		{
		 						case "SwitchResp":
		 							SwitchResp(Responsibility,parentWindow,waitframe,Facility);
		 				 			TestScenarioRow=k+1;
		 				 			continue;
		 				 	
		 						case "SwitchUser":
		 							Set<String> AllWindowHandles = driver.getWindowHandles();
		 				           for(String win:AllWindowHandles)
		 				           {
		 				                  driver.switchTo().window(win).close();
		 				           }
		 							driver = new InternetExplorerDriver();
		 							loginPage(Responsibility,parentWindow,wait,Username,Password,Facility);
		 							TestScenarioRow=k+1;
		 				 			continue;
		 						case "BrowserClose":
		 							driver.quit();
		 							TestScenarioRow=k+1;
		 				 			continue;
		 						
		 				 		
		 						case "ApplicationClose":
		 				 		
		 				 				driver.switchTo().defaultContent();
		 				 			    waitframe.until(ExpectedConditions.numberOfWindowsToBe(2));
		 				 			    Set<String> wind=driver.getWindowHandles();
		 				 			    parentwindow(parentWindow);
		 				 			    driver.switchTo().frame("menucontent");
		 				 			    List<WebElement>Switchresp=(new WebDriverWait(driver,Duration.ofSeconds(3))).until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.tagName("img")));
		 				 			    Switchresp.get(0).click();
		 				 			    Switchresp.get(0).click();
		 				 			  TestScenarioRow=k+1;
			 				 			continue;
		 				 		
		 				 		case "ApplicationLogin":
		 				 		
		 				 			loginPage(Responsibility,parentWindow,wait,Username,Password,Facility);
		 				 			TestScenarioRow=k+1;
		 				 			continue;
		 						}
		 		for(int n=1;n<=TestcaseTotalRow-1;n++)
			 		{
		 			try{
		 				
						if (TestCaseSheetIN.getRow(n).getCell(0).getStringCellValue()!= null && TestCaseSheetIN.getRow(n).getCell(0).getStringCellValue().length() != 0 ) 
						{ 
							Module=TestCaseSheetIN.getRow(n).getCell(0).getStringCellValue();
							
						}
						}
						catch(Exception e){}
		 				int Counts = wincount(n,TestCaseSheetIN);
 					int winnum = Counts;
 					
 					try{
		 				if(Count==0 && !TestCaseSheetIN.getRow(n).getCell(2).getStringCellValue().equals("Static") && !Module.contains("Wait")&& !Module.contains("TabOut") && !Module.contains("winclose")&& !Module.contains("ESCKey"))
	 					{
						try{
		 				if(TestCaseSheetIN.getRow(n).getCell(2).getStringCellValue()!= null && TestCaseSheetIN.getRow(n).getCell(4).getStringCellValue().equals("eHospital Information System"))
					{
						//Window_Frame_Handling.currentwindow(driver,"eHospital information system",wait,2,Module);
		 					Window_Frame_Handling.switchToNewWindow(wait,2,driver,"eHospital information system",Module);

					}
					}
					catch(Exception e)
		 			{
						title=Window_Frame_Handling.window(waitwin,3,driver,"iSOFT EM ver");
		 				
					}
					}
		 			}
		 			
		 				catch(Exception e1)
		 				{}
 					try{
 						if (TestCaseSheetIN.getRow(n).getCell(1).getStringCellValue()!= null &&  TestCaseSheetIN.getRow(n).getCell(0).getStringCellValue().length() != 0) {
 							ModFunc=TestCaseSheetIN.getRow(n).getCell(1).getStringCellValue();
 							}
 						}
 						catch(NullPointerException e){}
 					
 					CaptureDynamicData.cap_Dyn_Data(Module,driver,ID,DataSelection,TestCaseWorkbookIN,TestCaseExcelIN,evaluator);	
 					
					
					int p1=2;
					
				for(p=p1;p<=40;p++)
				
				{
					try{
						Object1=null;
						ProName1=null;
						ProValue1=null;
						AddProname=null;
						AddProValue=null;
						
						try{
						if(Flag.equals("True"))
						{
							Parameter1=null;
			 				p=p+4;
			 				break;
						}
						else
						{
							continue;
						}
						}
						catch(Exception e){}
						try{
						
							 if(Module.equals("winclose"))
			 				 {
			 					 try{
			 					String var = driver.getWindowHandle();
			 					String vartitle=driver.switchTo().window(var).getTitle();
			 					
			 					driver.close();
			 					Module=null;
			 					break;
			 					 }
			 					 catch(Exception e)
			 					 {
			 						driver.switchTo().activeElement();
				 					driver.close();
				 					Module=null;
				 					break;
				 					
			 					 }
			 					
			 				 }
				 			else if(Module.equals("winclose") && TestCaseSheetIN.getRow(n).getCell(7).getStringCellValue()!=null)
				 			{
				 				String win=TestCaseSheetIN.getRow(n).getCell(9).getStringCellValue();
				 				String windownum=TestCaseSheetIN.getRow(n).getCell(1).getStringCellValue();
				 				int windownumber=Integer.parseInt(windownum);
				 				Window_Frame_Handling.switchToNewWindow(wait,windownumber,driver,win,Module);
				 				System.out.println(driver.getTitle());
				 				driver.close();
				 				Module=null;
				 				break;
				 			}
				 		
							
				 			}
				 			catch(Exception e){}
						
						
						Parameter1=Data_Selection.cellType(p,n,TestCaseSheetIN,TestCaseWorkbookIN);
						if(Parameter1!=null && !Parameter1.isEmpty())
						{
							
							Object1=Parameter1;
						}
						try{
							if (TestCaseSheetIN.getRow(n).getCell(2).getStringCellValue().equals(null) && TestCaseSheetIN.getRow(n).getCell(2).getStringCellValue().length() == 0 && !TestCaseSheetIN.getRow(n).getCell(7).getStringCellValue().equals(null))
							{ 
								break;
							}
							}
							catch(Exception e)
			 				{
								break;
			 				}
							p2 = p+1;
						
							Parameter1=Data_Selection.cellType(p2,n,TestCaseSheetIN,TestCaseWorkbookIN);
						if(Parameter1!=null && !Parameter1.isEmpty())
						{
							
							ProName1=Parameter1;
						}
							p3 = p2+1;
							Parameter1=Data_Selection.cellType(p3,n,TestCaseSheetIN,TestCaseWorkbookIN);
							if(Parameter1.contains(".") && Parameter1.contains("E"))
							{
								Parameter1=Parameter1.replace(".", "");
								try{
								Parameter1=Parameter1.substring(0, 12);
								}
								catch(Exception e)
								{
									Parameter1=Parameter1.substring(0, 8);
								}
							}
						if(Parameter1!=null && !Parameter1.isEmpty())
						{
							
							ProValue1=Parameter1;
						}
							p4 = p3+1;
					
							Parameter1=Data_Selection.cellType(p4,n,TestCaseSheetIN,TestCaseWorkbookIN);
						if(Parameter1!=null && !Parameter1.isEmpty())
						{
							
							AddProname=Parameter1;
						}
							p5 = p4+1;
						
							Parameter1=Data_Selection.cellType(p5,n,TestCaseSheetIN,TestCaseWorkbookIN);
						if(Parameter1!=null && !Parameter1.isEmpty())
						{
							AddProValue=Parameter1;
						}
						
					}
					catch(NullPointerException e)
					{
						break;
					}
				try{					
					if(ProName1.equalsIgnoreCase("attribute/name"))
	 				 {
	 					ProName1="name"; 
	 				 }
					
					if((!Object1.isEmpty())&&(!ProName1.isEmpty())&&(!ProValue1.isEmpty()))
	 				 {
						driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
//						Thread.sleep(10000);
	 				 System.out.println("Object1....."+Object1);
	 				 System.out.println("Object1....."+ProName1);
	 				 System.out.println("Object1....."+ProValue1);
	 				if(Object1.equalsIgnoreCase("Browser") && !(ProValue1.equalsIgnoreCase("eHospital Information System")))
	 				 {
									 try{
	 				
	 						winnum=Counts+1;
		 					
				
	 				 }
	 					 catch(Exception e)
	 					 {
	 						 Set<String>winh=driver.getWindowHandles();
	 						 for(String win:winh)
	 						 {
	 							 driver.switchTo().window(win);
	 						 }
	 						 
	 					 }
	 				 }
					
	 				 else if(Object1.equalsIgnoreCase("window"))
	 				 {
	 					 System.out.println("window start here *****************"+Starttime());
	 					winnum=winnum+2;
	 					System.out.println("Counts                                                   "+Counts);
	 					if(Count==1)
	 					{//do nothing 
	 					}
	 					else if(Count==2)
	 					{
	 						try
	 						{
	 							Parameter1=Data_Selection.cellType(14,n,TestCaseSheetIN,TestCaseWorkbookIN);
	 							ProValue1=Parameter1;
	 						}
	 						catch(Exception e)
	 						{	}
	 						p=p+10;
	 						
	 					}
	 					else if(Count==3)
	 					{
	 						try
	 						{
	 							Parameter1=Data_Selection.cellType(19,n,TestCaseSheetIN,TestCaseWorkbookIN);
	 							ProValue1=Parameter1;
	 							p=p+10;
	 						}
	 						catch(Exception e)
	 						{}
	 					}
	 					else if(Count==4)
                       {
                              try
                              {
                                     Parameter1=Data_Selection.cellType(24,n,TestCaseSheetIN,TestCaseWorkbookIN);
                                     ProValue1=Parameter1;
                                     p=p+15;
                              }
                              catch(Exception e)
                              {     
                              }
                               }
	 					try
	 						{
	 							Window_Frame_Handling.currentwindow(driver,ProValue1,wait,winnum,Module);
	 							if(!driver.getTitle().contains(ProValue1))
	 							{
	 								Window_Frame_Handling.switchToNewWindow(wait,winnum,driver,ProValue1,Module);
	 							}
	 						}
	 						catch(Exception e)
	 						{
	 							Window_Frame_Handling.switchToNewWindow(wait,winnum,driver,ProValue1,Module);
	 						}
	 					 System.out.println("window end here *****************"+Endtime());
                      	 					 }

	 					 else if(Object1.equalsIgnoreCase("Frame"))
	 				  {
	 						 System.out.println("Frame starts here *****************"+Starttime());
	 					  try{
	 						 Parameter1=Data_Selection.DataSelection(DataSelection,TestCaseSheetIN,n,p,TestCaseWorkbookIN);
	 						 framename=Window_Frame_Handling.framehandle(ProValue1,driver,Parameter1);
	 					  	}

	 					  catch(Exception e)
	 					  { }
	 					 System.out.println("Frame end here *****************"+Endtime());
	 				    }
                      else if(Object1.equalsIgnoreCase("RTEditor"))
                      {
                             try{
                                   Parameter1=Data_Selection.DataSelection(DataSelection,TestCaseSheetIN,n,p,TestCaseWorkbookIN);
                                   
                                   if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
           	 					{
           	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
           	 					}
           	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
           	 					{
           	 						CheckPoint.checkpoint(filePath,workbook,ProName1,Object1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
           	 					}
                                  }

                             catch(Exception e)
                             { }
                        }
                      else if(Object1.equalsIgnoreCase("PDFVerfication"))
                      {
                             try{
                            	 Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
                                   
                                   if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
           	 					{
                                	   CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
           	 					}
                                   else if(Module.equalsIgnoreCase("NegativeCheckpoint") || ModFunc.equalsIgnoreCase("NegativeCheckpoint"))
              	 					{
                                	   CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
              	 					}
           	 					 }

                             catch(Exception e)
                             { }
                        }
	 				 else if(Object1.equalsIgnoreCase("Windowactivate"))
	 				  {
	 					  try{
	 						 Parameter1=Data_Selection.DataSelection(DataSelection,TestCaseSheetIN,n,p,TestCaseWorkbookIN);
	 						CaptureDynamicData.cap_Dyn_Data(ProValue1,driver,ID,DataSelection,TestCaseWorkbookIN,TestCaseExcelIN,evaluator);
	 					  	}

	 					  catch(Exception e)
	 					  { }
	 				    }
					
	 				 else if(Object1.equalsIgnoreCase("WebElement"))
	 				  {
	 					 try
	 					 {
	 						 try{
	 						Parameter1=Data_Selection.DataSelection(DataSelection,TestCaseSheetIN,n,p,TestCaseWorkbookIN);
	 						 }
	 								catch(Exception e){}
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 						Link(filePath,workbook,ProName1,ProValue1,wait,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue);
	 					}
	 					 }
	 					 catch(Exception e)
	 					 {
	 						 result="FAIL";
	 						 comments="WebElement not Clicked";
		 					ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					 }
	 					
	 				  }
					
	 				  else if(Object1.equalsIgnoreCase("link"))
	 				  {
	 					  
	 					  try{
	 							Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);	
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{	
	 						if(ProName1.equalsIgnoreCase("name")){ProName1="linktext";}
	 							
	 						Link(filePath,workbook,ProName1,ProValue1,wait,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue);
	 						
	 						}
	 					if(Module.equalsIgnoreCase("StaticCheckPoint"))
                        { 
                                      Object1="Static";
                                      ProValue1=Data_Selection.cellType(4,n,TestCaseSheetIN,TestCaseWorkbookIN);
                                      CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
                                      isAlertPresent(driver);
                               }
	 					  }
	 					  catch(Exception e)
	 					  {
	 						 result="FAIL";
	 						 comments="Link not Clicked"+ "    "+e;
	 						 ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					  }
	 					 try
							{
			 				isAlertPresent(TestCaseWorkbookIN, TestCaseExcelIN);
							}
								catch(Exception e)
								{
									result="FAIL";
			 						comments="WebButton not clicked";
			 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
								}
	 				  }
	 				 else if(Object1.equalsIgnoreCase("Image"))
	 				  {
	 					 
	 					 try{
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);	
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 					//Image(filePath,workbook,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue);
	 						Link(filePath,workbook,ProName1,ProValue1,wait,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue);
	 					}
	 				  }
	 					catch(Exception e)
	 					 {
	 						
	 						result="FAIL";
	 						comments="Image not Clicked";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					 }
	 				  }
	 				
					
	 				 else if(Object1.equalsIgnoreCase("WebCheckBox"))
	 				  {
	 					 try{
	 						Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
	 					 
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint")|| ModFunc.equalsIgnoreCase("CheckedCheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 						
	 					WebCheckBox(filePath,workbook,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue,Parameter1);
	 					}

	 					 }
	 					 catch(Exception e)
	 					 {
	 						result="FAIL";
	 						comments="WebCheckBox not Clicked";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					 }
	 					try
						{
		 				isAlertPresent(TestCaseWorkbookIN, TestCaseExcelIN);
						}
							catch(Exception e)
							{
								result="FAIL";
		 						comments="WebButton not clicked";
		 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
							}
	 				  }
					
	 				else if(Object1.equalsIgnoreCase("WebRadioGroup"))
	 				  {
	 					try{
	 						Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 						Link(filePath,workbook,ProName1,ProValue1,wait,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue);
	 					}
	 				
	 					}
	 					catch(Exception e)
	 					{
	 						result="FAIL";
	 						comments="WebRadioGroup not Selected";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					}
	 				  }
	 				else if(Object1.equalsIgnoreCase("Static"))
	 				  {
	 					try{
	 						Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					}
	 					catch(Exception e)
	 					{
	 						result="FAIL";
	 						comments="Static not Selected";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					}
	 				  }
					
					 else if(Object1.equalsIgnoreCase("WebEdit"))
	 				  {
	 					try{
	 						
	 						Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
	 						try{
	 						if(Parameter1.contains("E") && !Parameter1.equals("denture") && !Parameter1.equals("ENT") && !Parameter1.equals("ENT General") && !Parameter1.equals("Examination") && !driver.getTitle().equals("Surgery Type") && !TestCaseSheetIN.getRow(n).getCell(p).getCellFormula().equalsIgnoreCase("Menu!B65536"))
	 						{ 
	 							Parameter1=Parameter1.replace("E", "");
	 						}
	 					}
 						catch(Exception e){}
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 					WebEdit(filePath,workbook,ProName1,ProValue1,Parameter1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue,n);
	 					}
	 					if(Module.equalsIgnoreCase("StaticCheckPoint"))
                        { 
                                       Object1="Static";
                                      ProValue1=Data_Selection.cellType(4,n,TestCaseSheetIN,TestCaseWorkbookIN);
                                      CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
                                      isAlertPresent(driver);
                               }
	 					}
	 					catch(Exception e)
	 					{
	 						result="FAIL";
	 						comments="WebEdit not Entered";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					}
	 				 }
					
	 				else if(Object1.equalsIgnoreCase("WebList"))
	 				  {
	 					try{
	 						Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
	 					if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);	
	 					}
	 					else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
	 						CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 					WebList(filePath,workbook,ProName1,ProValue1,Parameter1,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					}
	 				
	 					}
	 					catch(Exception e)
	 					{
	 						result="FAIL";
	 						comments="WebList not Selected";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
	 					}
	 					if(Module.equalsIgnoreCase("StaticCheckPoint"))
                        { 
                                       Object1="Static";
                                      ProValue1=Data_Selection.cellType(4,n,TestCaseSheetIN,TestCaseWorkbookIN);
                                      CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
                                      isAlertPresent(driver);
                               }

	 				  }
					
	 				else if(Object1.equalsIgnoreCase("WinButton"))
	 				 {
	 					isAlertPresent(TestCaseWorkbookIN,TestCaseExcelIN);  
	 					/*driver.switchTo().alert();
	 					  driver.findElement(By.name("OK")).click();
	 					  isAlertPresent(driver);*/
	 					  break;
	 					
	 				 }
					else if(Object1.equalsIgnoreCase("WebButton"))
	 				 {
						try{
						Parameter1=Data_Selection.ModDataSel(Module,TestCaseSheetIN,n,p,TestCaseWorkbookIN,DataSelection);
						if(Module.equalsIgnoreCase("CheckPoint") || ModFunc.equalsIgnoreCase("CheckPoint"))
	 					{
							CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);	
	 					}
						else if(Module.equalsIgnoreCase("NegativeCheckpoint"))
	 					{
							CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
	 					}
	 					else
	 					{
	 						Link(filePath,workbook,ProName1,ProValue1,wait,driver,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow,AddProValue);
	 						}
						if(Module.equalsIgnoreCase("StaticCheckPoint"))
                        { 
                                       Object1="Static";
                                      ProValue1=Data_Selection.cellType(4,n,TestCaseSheetIN,TestCaseWorkbookIN);
                                      CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,waitele,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
                               isAlertPresent(driver);
                               Module=null;
                               }

						}
						catch(Exception e)
						{
							result="FAIL";
							System.out.println(e);
	 						comments="WebButton not Clicked";
	 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
						}
						try
						{
		 				isAlertPresent(TestCaseWorkbookIN, TestCaseExcelIN);
						}
							catch(Exception e)
							{
								result="FAIL";
		 						comments="WebButton not clicked";
		 						ResultReporting.printresult(filePath,workbook,Step_Report,CheckPointsh,Module,ModFunc,result,comments,TestScenarioRow);
							}
	 				 }
	 	 //********************End Code For objects Link,WebCheckbox,radiobutton,button************************************
	 				Parameter1=null;
	 				p=p+4;
	 				
	 				 }
					
	 				else
	 				{
	 					break;
	 				}
					}
					catch(Exception e){}
	 			
	 				   }
				try{
				if (TestCaseSheetIN.getRow(n+1).getCell(0).getStringCellValue().equals(null) && TestCaseSheetIN.getRow(n+1).getCell(0).getStringCellValue().length() == 0 ) 
				{ 
					break;
					
				}
				}
				catch(Exception e)
 				{
					break;
 				}
				 		}
	 		TestScenarioRow=k+1;
	 		End_Time=Endtime();
	 		Elapsed_Time_in_seconds=Elapsedtime(Start_Time, End_Time);
	 		ResultReporting.printresulttime(filePath,workbook,Exectimesh,TestScenario,Start_Time,End_Time,Elapsed_Time_in_seconds,TestScenarioRow);
	 			 		}
	 		System.exit(0);
	}
 	 catch(Exception e)
		 				{System.out.println(e);}
 
	}
	
	

	@SuppressWarnings("unused")
	
	public static void SwitchResp(String responsibility,String parentWindow,WebDriverWait waitframe,String Facility) throws InterruptedException
    {
		try{
			driver.switchTo().defaultContent();
		    waitframe.until(ExpectedConditions.numberOfWindowsToBe(2));
		    Set<String> wind=driver.getWindowHandles();
		    Thread.sleep(1000);
		   for(String w:driver.getWindowHandles())
		   {
			   String title=driver.switchTo().window(w).getTitle();
				if(title.equals("eHospital Information System"))
				{
					break;
				}
		   }
		    System.out.println(driver.getTitle());
		    driver.switchTo().frame("menucontent");
		   WebElement switchele=driver.findElement(By.xpath("//img[@title='Switch Responsibility']"));
		   JavascriptExecutor jsExecutor = (JavascriptExecutor)driver;
		   jsExecutor.executeScript("var elem=arguments[0]; setTimeout(function() {elem.click();}, 100)",switchele);
		    Thread.sleep(1000);
		    waitframe.until(ExpectedConditions.numberOfWindowsToBe(2));
		    Set<String> wind1=driver.getWindowHandles();
		    for(String win:wind1)
		{
			String title=driver.switchTo().window(win).getTitle();
			if(title.equals("Switch Responsibility"))
			{
				System.out.println("Switch Responsibility title********************"+title);
				break;
			}
		}
		driver.findElement(By.xpath("//input[@class='dhx_combo_input']")).sendKeys(responsibility);
		driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.DOWN);
		driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.ENTER);
		driver.findElement(By.id("changeOK")).click();
		Thread.sleep(2000);
		}
		catch(Exception e){System.out.println(e);}
		  }
		
	public static String parentwindow(String parentWindow)
	{
		parentWindow = driver.getWindowHandle();
		System.out.println(driver.getTitle());
		return parentWindow;
	}
	
	public static void loginPage(String Responsibility,String parentWindow,WebDriverWait wait,String Username,String Password,String Facility) throws InterruptedException
	{
		try{
			
			driver.navigate().to("http://130.78.63.1:8989/HIS/eSM/jsp/login.jsp?computerName=%22DUC-MVa1lIs6W0u%22&ipAddress=%22192.168.0.203%22");//Chrome
//			driver.navigate().to("http://130.78.63.42:9999/HIS/eSM/jsp/login.jsp");//em12xsel
			//driver.navigate().to("http://cscdbche754:9999/HIS/eSM/jsp/login.jsp");//em12xsut
//			driver.navigate().to("http://130.78.62.220:9999/HIS/eSM/jsp/login.jsp");//MOHML
			//driver.navigate().to("http://cscdbche754:7777/HIS/eSM/jsp/login.jsp");//MOHML
			//driver.navigate().to("http://cscindae753150:7001/eHIS/eCommon/jsp/home.jsp");//Sikirin
			//System.out.println("�rror");
	driver.findElement(By.name("name")).sendKeys(Username);
	driver.findElement(By.name("password")).sendKeys(Password);
				WebElement loginbtn = (new WebDriverWait(driver,Duration.ofSeconds(5)))
  		  .until(ExpectedConditions.presenceOfElementLocated(By.id("loginID")));
	loginbtn.click();
	parentwindow(parentWindow);
	driver.findElement(By.className("dhx_combo_input")).sendKeys(Responsibility);
	driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.DOWN);
	driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.ENTER);
	driver.findElement(By.className("dhx_combo_input")).sendKeys(Keys.TAB);
	driver.findElement(By.id("loginID")).click();
//	driver.navigate().to("http://130.78.62.220:9999/HIS/eCommon/jsp/eHIS.jsp");   //only for edgeuse
		}
		catch(Exception e){System.out.println(e);}
	}

	@SuppressWarnings("deprecation")
	public static boolean isAlertPresent(HSSFWorkbook TestCaseWorkbookIN,FileInputStream TestCaseExcelIN){
        try{
        	String CapDataFile=ConstVariable.CapDataFile;
        	WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(10));
        	if(Module.equalsIgnoreCase("OA_Appt_ID"))
        	{
        		
        		Alert alert = driver.switchTo().alert();
        		String text=alert.getText();
        		String[] Appt_ID=text.split("No.");
        		System.out.println("Appt_ID     "+Appt_ID[1]);
        		CaptureDynamicData.SetDataToExcel("OA_Appt_ID",9,1,TestCaseWorkbookIN,TestCaseExcelIN,CapDataFile);
        		CaptureDynamicData.SetDataToExcel(Appt_ID[1],9,1,TestCaseWorkbookIN,TestCaseExcelIN,CapDataFile);
        		CaptureDynamicData.evaluator(TestCaseWorkbookIN,evaluator,CapDataFile);
				HSSFCell cell = TestCaseWorkbookIN.getSheetAt(0).getRow(65535).getCell(11);
				evaluator.evaluateFormulaCell(cell);
				}
        	else if(Module.equalsIgnoreCase("MR_Request_No"))
        	{
        		String text=driver.switchTo().alert().getText();
        		String[] Req_ID=text.split("ID :");
        		System.out.println("Req_ID     "+Req_ID[1]);
        		CaptureDynamicData.SetDataToExcel(Req_ID[1],6,0,TestCaseWorkbookIN,TestCaseExcelIN,CapDataFile);
        		CaptureDynamicData.evaluator(TestCaseWorkbookIN,evaluator,CapDataFile);
				HSSFCell cell = TestCaseWorkbookIN.getSheetAt(0).getRow(65535).getCell(49);
				evaluator.evaluateFormulaCell(cell);
				}
        	
        	else if(Module.equalsIgnoreCase("OA_Appt_ID_WO_PatID"))
        	{
        		String text=driver.switchTo().alert().getText();
        		String[] OA_Appt_ID_WO_PatID=text.split(". No.");
        		System.out.println("OA_Appt_ID_WO_PatID     "+OA_Appt_ID_WO_PatID[1]);
        		CaptureDynamicData.SetDataToExcel(OA_Appt_ID_WO_PatID[1],6,0,TestCaseWorkbookIN,TestCaseExcelIN,CapDataFile);
        		CaptureDynamicData.evaluator(TestCaseWorkbookIN,evaluator,CapDataFile);
				HSSFCell cell = TestCaseWorkbookIN.getSheetAt(0).getRow(65535).getCell(11);
				evaluator.evaluateFormulaCell(cell);
				}
        else if(Module.equalsIgnoreCase("RTVDialog"))
        	{
        		try{
        			Thread.sleep(2000);
        		 wait.until(ExpectedConditions.alertIsPresent());
        		Alert alert = driver.switchTo().alert();
        		alert.accept();
        		
           		}
        		catch(Exception e)
        		{
        			 Actions act=new Actions(driver);
            		 act.sendKeys(Keys.RETURN);
            		 act.sendKeys(Keys.ENTER);
            	}
        					
        	}
        else if(Module.equalsIgnoreCase("RTVDialogCancel"))
    	{
    		try{
    			Thread.sleep(2000);
    		 wait.until(ExpectedConditions.alertIsPresent());
    		Alert alert = driver.switchTo().alert();
    		alert.accept();
    		alert.dismiss();
    		
       		}
    		catch(Exception e)
    		{
    			 Actions act=new Actions(driver);
        		 act.sendKeys(Keys.RETURN);
        		 act.sendKeys(Keys.ENTER);
        	}
    					
    	}
        else if(Module.equalsIgnoreCase("RTVDialog2"))
    	{
    		try{
    			Thread.sleep(2000);
    		 wait.until(ExpectedConditions.alertIsPresent());
    		Alert alert = driver.switchTo().alert();
    		alert.accept();
    		
       		}
    		catch(Exception e)
    		{
    			 Actions act=new Actions(driver);
        		 act.sendKeys(Keys.RETURN);
        		 act.sendKeys(Keys.ENTER);
        	}
    		try{
    			Thread.sleep(2000);
    		 wait.until(ExpectedConditions.alertIsPresent());
    		Alert alert = driver.switchTo().alert();
    		alert.accept();
    		
       		}
    		catch(Exception e)
    		{
    			 Actions act=new Actions(driver);
        		 act.sendKeys(Keys.RETURN);
        		 act.sendKeys(Keys.ENTER);
        	}
    					
    	}
        	
        	else if(Module.equalsIgnoreCase("ReschApp"))
        	{
        		driver.getWindowHandle();
        		Set<String>winh=driver.getWindowHandles();
        		for(String win:winh)
        		{
        			driver.switchTo().window(win);
        		}
        		
        		try{
        		Thread.sleep(1000);
        		//WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(5));
        		wait.until(ExpectedConditions.alertIsPresent());
        		Alert alert = driver.switchTo().alert();
        		alert.accept();
           		driver.getWindowHandle();
        		}
        		catch(Exception e){}
        		}
        	else if(ProValue1.equalsIgnoreCase("Cancel"))
        	{ driver.switchTo().alert().dismiss();
        		}
        	else if(ProValue1.equalsIgnoreCase("Ok"))
            {
           	Set <String>a=driver.getWindowHandles();
           	try{
           	for(String b:a){
           		
           		driver.switchTo().window(b);
           		try{
           			
           		 driver.switchTo().alert().accept();
           		}
           		catch(Exception e){continue;}
           		 break;
           	}
           	}
           	catch(UnhandledAlertException e){
           	
           	}
                   }
        	else if(Module.equalsIgnoreCase("StaticCheckPoint"))
        	{ 
        		CheckPoint.checkpoint(filePath,workbook,Object1,ProName1,ProValue1,wait,driver,Step_Report,CheckPointsh,Module,ModFunc,TestScenarioRow,Parameter1);
        		isAlertPresent(driver);
        		}
        	
        	else
        	{
        		JavascriptExecutor js = (JavascriptExecutor) driver;
				

        		js.executeScript("window.focus();");
        		 driver.switchTo().alert().accept();

        	}
        	 return true;
        }
        catch(Exception e){
        	return false;
        }
        finally
        {}
    }
	
		
	

	
	
	public static boolean isAlertPresent(WebDriver driver)
	{
        try{
            driver.switchTo().alert().accept();
            return true;
        }
        catch(Exception e){
            return false;
        }
        finally
        { }
        
	}
	

public static void WebEdit(String filePath,HSSFWorkbook workbook,String locatorType, String value,String Parameter1,WebDriverWait waitele,WebDriver driver,HSSFSheet step_Report,HSSFSheet checkPointsh,String module1,String modFunc1,String result,String comments,int TestScenarioRow,String AddProValue,int n) throws InterruptedException, WriteException, IOException 
{
	System.out.println("Parameter1******"+Parameter1);
	try {
		
		switch (locatorType){
		case "name":
			 ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.name(value)));
			break;
			
		case "id":
            ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.id(value)));
               
                  break;

		case "xpath":
			
				 ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.xpath(value)));
				
				break;
				
		case "css":
			
			 ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(value)));
			
			break;
			
		   case "clinicalnote":
	             JavascriptExecutor js = (JavascriptExecutor) driver;
	        WebElement ele3=driver.findElement(By.name("C_CLT0000000000021"));
	        js.executeScript("arguments[0].setAttribute('value','Tested');",ele3);
	            break;
	            
		   case "clinicalnote1":
	             JavascriptExecutor js1 = (JavascriptExecutor) driver;
	        WebElement ele4=driver.findElement(By.name("C_ALT0000000000651"));
	        js1.executeScript("arguments[0].setAttribute('value','Tested');",ele4);
	            break;
	            
		   case "clinicalnote2":
	             JavascriptExecutor js2 = (JavascriptExecutor) driver;
	        WebElement ele5=driver.findElement(By.name("C_CLT0000000000162"));
	        js2.executeScript("arguments[0].setAttribute('value','ACT');",ele5);
	            break;   
	            
		   case "clinicalnote3":
	             JavascriptExecutor js3 = (JavascriptExecutor) driver;
	        WebElement ele6=driver.findElement(By.name("C_ALT0000000000022"));
	        js3.executeScript("arguments[0].setAttribute('value','MedRept');",ele6);
	            break;   
		

		}
		if(ele.isDisplayed())
		{
			if(Parameter1.equals("Click"))
			{
				ele.click();
				WebElement body_element = driver.findElement(By.tagName("body"));
				body_element.click();
				body_element.sendKeys(Keys.TAB);
			}
			else
			{
			ele.clear();
			ele.sendKeys(Parameter1);
			result="PASS";
			comments="WebEdit Entered";
			ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
		}
		else
		{
			result="FAIL";
			comments="WebEdit not Entered";
			ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		}
	} 
	catch (NoSuchElementException e) {
		result="FAIL";
		comments="WebEdit not Entered";
		ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		System.err.format("No Element Found to click link" + e);
	}
	
}


	
public static void WebList(String filePath,HSSFWorkbook workbook,String locatorType, String value,String Parameter1,WebDriver driver,HSSFSheet step_Report,HSSFSheet checkPointsh,String module1,String modFunc1,String result,String comments,int TestScenarioRow) throws InterruptedException, WriteException, IOException 
{
	try {
		switch (locatorType){
		case "id":
			Select dropdown=new Select(driver.findElement(By.id(value)));
			dropdown.selectByVisibleText(Parameter1);
			break;
			
		case "name":
			 WebElement ele = new WebDriverWait(driver,Duration.ofSeconds(30))
	            .until(ExpectedConditions.refreshed(
	                    ExpectedConditions.presenceOfElementLocated(By.name(value))));
		
//			Thread.sleep(10000);
//			Select dropdown1=new Select(ele);
			if(ele.isDisplayed())
			{
				try{
//		System.out.println(Parameter1);
//				dropdown1.selectByVisibleText(Parameter1);
		Window_Frame_Handling.selectOption(driver, ele, Parameter1);	
//		dropdown1.selectByIndex(2);
				}
				
				catch(Exception e){System.out.println(e);}
//				catch(Exception e){dropdown1.selectByValue(Parameter1);}
				result="PASS";
				comments="WebList Selected";
				ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
			else
			{
				result="FAIL";
				comments="WebList not Selected";
				ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
			break;
			case "index":
			WebElement ele3 = (new WebDriverWait(driver,Duration.ofSeconds(5))).until(ExpectedConditions.presenceOfElementLocated(By.name(value)));
			Select dropdown4=new Select(ele3);
			if(ele3.isEnabled())
			{
				JavascriptExecutor js = (JavascriptExecutor) driver;
				try{
				js.executeScript("arguments[0].sendKeys(Parameter1);", ele3);
				}
				catch(Exception ee){
				dropdown4.selectByIndex(Integer.parseInt(Parameter1));
				}
				result="PASS";
				comments="WebList Selected";
				ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
			else
			{	
				result="FAIL";
				comments="WebList not Selected";
				ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
			break;
			
			
		case "xpath":
			
			List<WebElement> dropdown2=driver.findElements(By.xpath("//Select[@name='name_prefix']"));
			Select dropdown3=new Select(dropdown2.get(1));
			dropdown3.selectByVisibleText(Parameter1);
			result="PASS";
			comments="WebList Selected";
			ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			break;
			
	
			
	case "collect":
			List<WebElement> tagname=driver.findElements(By.tagName("select"));
			for(WebElement collecttagname:tagname)
			{
						Select options=new Select(collecttagname);
						options.selectByVisibleText(Parameter1);
						result="PASS";
						comments="WebList Selected";
						ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
						break;
					}
				
		default:
			System.out.println("WebList switchcase default access no matches found in PropertyName");
			break;
		}
		
	} catch (NoSuchElementException e) {
		result="FAIL";
		comments="WebList not Selected";
		ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		System.err.format("No Element Found to click link" + e);
	}
}
	
public static void Link(String filePath,HSSFWorkbook workbook,String locatorType, String value,WebDriverWait waitele,WebDriver driver,HSSFSheet step_Report,HSSFSheet checkPointsh,String module1,String modFunc1,String result,String comments,int TestScenarioRow,String AddProValue) throws InterruptedException, WriteException, IOException 
{
	WebElement ele = null;
	try{	
	switch (locatorType){
		case "name":
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.name(value)));
						
			break;
		case "linktext":
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.linkText(value)));
			break;
			
		case "id":
			 ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.id(value)));
			break;
		case "title":
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@title='"+value+"']")));
			break;
		case "value":
				ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@value='"+value+"']")));
				break;
		case "xpath":
		 ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.xpath(value)));
			break;
		case "css":
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(value)));
		break;
		case "class":
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.className(value)));
		break;

            
		default:
			System.out.println("Link switchcase default access no matches fount in PropertyName");
			}
		if(ele.isDisplayed())
			{		
				//ele.click();
				JavascriptExecutor js = (JavascriptExecutor) driver;
					js.executeScript("var elem=arguments[0]; setTimeout(function() {elem.click();}, 100)",ele);
				result="PASS";
				comments="Link Clicked";
				ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
			else
			{
				result="FAIL";
				comments="Link not Clicked";
				ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
			}
	}
	 catch (NoSuchElementException e) {
		result="FAIL";
		comments="Link not Clicked";
		ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		System.err.format("No Element Found to click link" + e);
		
	}
}

		
public static void WebCheckBox(String filePath,HSSFWorkbook workbook,String locatorType,String ProValue1,WebDriverWait waitele,WebDriver driver,HSSFSheet step_Report,HSSFSheet checkPointsh,String module1,String modFunc1,String result,String comments,int TestScenarioRow,String AddProValue,String Parameter1) throws InterruptedException, WriteException, IOException, AWTException 
{
	try {
		
		switch (locatorType){
		case "name":
		
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.name(ProValue1)));
			
			break;
			
		case "id":
			
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.id(ProValue1)));
			
			break;

		case "xpath":
			JavascriptExecutor js1 = (JavascriptExecutor) driver;
			ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.xpath(ProValue1)));
			
			break;
			
		//case "xpath":
			
			// ele=waitele.until(ExpectedConditions.presenceOfElementLocated(By.xpath(value)));
			
			//break;
		default:
			System.out.println("WebCheckBox switchcase default access no matches fount in PropertyName");
			break;
		}
		if(ele.isDisplayed())
		{
			try{
			if(Parameter1.equals("ON"))
			{
				if (ele.isSelected())
				{
					
				}
				else
				{
					ele.click();
				}
			}
			else if(Parameter1.equals("OFF"))
			{
				if (ele.isSelected())
				{
					ele.click();
				}
				else
				{
					
				}
			}
			}
			catch(Exception e){
			{
				if (ele.isSelected())
				{				
				}
				else
				{
					JavascriptExecutor js = (JavascriptExecutor) driver;
					js.executeScript("arguments[0].setAttribute('checked','true');",ele);
					
				}
			}
			}
			result="PASS";
			comments="WebCheckBox Clicked";
			ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		}
		else
		{
			result="FAIL";
			comments="WebCheckBox not Clicked";
			ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		}
		
	} catch (NoSuchElementException e) {
		result="FAIL";
		comments="WebCheckBox not Clicked";
		ResultReporting.printresult(filePath,workbook,step_Report,checkPointsh,module1,modFunc1,result,comments,TestScenarioRow);
		System.err.format("No Element Found to click link" + e);
		
	}
}
public static int wincount(int n,HSSFSheet TestCaseSheetIN)
{
Count = 0;

for(int s=2;s<=35;s++)
{
String checkwindow="Window";
HSSFCell cellvalue=TestCaseSheetIN.getRow(n).getCell(s);

try{
if(cellvalue.getStringCellValue().equalsIgnoreCase(checkwindow)&& cellvalue.getStringCellValue()!=null)
{
	Count=Count+1;
}

}
catch(Exception e){}
}
return Count;
}

public static String Starttime()
{

	SimpleDateFormat dateFormat = new SimpleDateFormat("hh:mm:ss aa");
	String Starttime = dateFormat.format(new Date()).toString();
	System.out.println(Starttime);
	return Starttime;
}


public static String Endtime() throws InterruptedException
{
	SimpleDateFormat dateFormat = new SimpleDateFormat("hh:mm:ss aa");
	
	String Endtime = dateFormat.format(new Date()).toString();
	System.out.println(Endtime);
	return Endtime;
}

public static long Elapsedtime(String Starttime,String Endtime) throws InterruptedException, ParseException
{
String time1 =Starttime;
String time2 =Endtime;

SimpleDateFormat format = new SimpleDateFormat("HH:mm:ss");
Date date1 = format.parse(time1);
Date date2 = format.parse(time2);
long Elapsedtime = date2.getTime() - date1.getTime();
Elapsedtime = Elapsedtime/1000;
System.out.println("Starttime    "+Starttime);
System.out.println("Endtime      "+Endtime);
System.out.println("Elapsedtime  "+Elapsedtime);
return Elapsedtime;
}

public static void scrolldown()
{
	try{
	((JavascriptExecutor) driver).executeScript("scroll(0,250);");
	}
	catch(Exception e)
	{
		((JavascriptExecutor) driver).executeScript("window.scrollBy(0,250)", "");
	}
	
}


}


