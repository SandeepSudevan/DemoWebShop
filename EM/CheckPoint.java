package EM;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;
import java.sql.Array;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.Spliterator;
import java.util.concurrent.TimeUnit;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.net.ssl.HttpsURLConnection;

import org.apache.http.HttpResponse;
import org.apache.http.StatusLine;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpHead;
import org.apache.http.client.methods.HttpUriRequest;
import org.apache.http.impl.client.CloseableHttpClient;
//import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.client.HttpClients;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

/*Description: This class File does the Validation/Verfication functionality.It does both Positive and Negative Verification*/

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import jxl.write.WriteException;

public class CheckPoint {
	@SuppressWarnings("null")
	
	public static void checkpoint(String filePath,HSSFWorkbook workbook,String Object1,String Propname,String ProValue1,WebDriverWait wait,WebDriver driver,HSSFSheet step_Report,HSSFSheet checkPointsh,String module1,String modFunc1,int TestScenarioRow,String Parameter1) throws WriteException, IOException
	{
		String result1 = null;
		String result2 = null;
		if(module1.equalsIgnoreCase("Checkpoint")||modFunc1.equalsIgnoreCase("Checkpoint"))
		{
			 result1="Pass";
			 result2="Fail";
		}
		else
		{
			 result1="Fail";
			 result2="Pass";
			 wait=new WebDriverWait(driver,Duration.ofSeconds(1));
		}
		WebElement xpath = null;
		try
		{
		switch(Object1)
		{
		
		case "Link":
		case "WebElement":
		case "Image":
				switch(Propname){
				case "name":
					xpath=driver.findElement(By.name(ProValue1));
					break;
				case "id":
					xpath=driver.findElement(By.id(ProValue1));
					break;
				case "xpath":
					xpath=driver.findElement(By.xpath(ProValue1));
					break;
		
					}
			
				if(Parameter1==null)
				{
					
		if(xpath.isDisplayed())
		{
			ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,Object1 +" " +ProValue1 +" Exist",TestScenarioRow);
			System.out.println(Object1+ " ***** "+result1);
		}
		
		else
		{
			ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1 +" " +ProValue1 +" does not Exist",TestScenarioRow);
			System.out.println(Object1+ " ***** "+result2);
		}
		}
 
			else
			{	
				System.out.println(xpath.getText());
				if(xpath.getText().contains(Parameter1))
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,Object1 +" " +Parameter1+" Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
					break;
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1 +" " +Parameter1+" does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
				}
				
			}
		break;
		case "PDFVerfication":
		if(Parameter1==null || Parameter1.equals(""))
		{
			//do nothing
		}
		else
		{
		 wait.until(ExpectedConditions.numberOfWindowsToBe(3));
		 System.out.println("size........"+driver.getWindowHandles().size());
		 for(String win:driver.getWindowHandles())
		 {        	 
				driver.switchTo().window(win);
				System.out.println("title......................"+driver.getTitle());
				System.out.println("currenturl......................"+driver.getCurrentUrl());
					if(driver.getCurrentUrl().contains("report"))
				{
					break;
				}
					else
				{
					continue;
				}
		 }
			 //driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			 String Currentlink=driver.getCurrentUrl();
			URL url=new URL(Currentlink);
			InputStream is=url.openStream();
			BufferedInputStream fp=new BufferedInputStream(is);
			PDDocument document=null;
			document=PDDocument.load(fp);
			PDFTextStripper stripper = new PDFTextStripper();
			String extractedText= stripper.getText(document);
			
			extractedText = extractedText.replace("\n", "").replace("\r", "");
			System.out.println(extractedText);
			if(extractedText.contains(Parameter1))
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"Text Exist in PDF File",TestScenarioRow);
					System.out.println("Text Exist in PDF File ***** "+result1);
					break;
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"Text does not Exist in PDF File",TestScenarioRow);
					System.out.println("Text does not Exist in PDF File ***** "+result2);
					 
				}
		}
		break;
		case "WebButton":
			switch(Propname){
			case "name":
				xpath=driver.findElement(By.name(ProValue1));
				break;	
		case "xpath":
				xpath=driver.findElement(By.xpath(ProValue1));
				break;
						}
			if(Parameter1==null)
			{
			if(xpath.isDisplayed())
			{
				ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebButton Exist",TestScenarioRow);
				System.out.println(Object1+ " ***** "+result1);
			}
			else
			{
				ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebButton does not Exist",TestScenarioRow);
				System.out.println(Object1+ " ***** "+result2);
			}
			}
			else
			{
				String disabled=xpath.getAttribute("disabled");
				String value=xpath.getAttribute("value");
				boolean display=xpath.isDisplayed();
				boolean enable=xpath.isEnabled();
				if(Parameter1.equalsIgnoreCase(disabled)||Parameter1.equalsIgnoreCase(value)||Parameter1.equals(display)||Parameter1.equals(enable))
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebButton Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebButton does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
				}
					
              }
        
        break;

		case "WebTable":
			xpath=driver.findElement(By.xpath(ProValue1));
			
			if(xpath.isDisplayed())
			{
				ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebTable Exist",TestScenarioRow);
				System.out.println(Object1+ " ***** "+result1);
			}
			else
			{
				ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebTable does not Exist",TestScenarioRow);
				System.out.println(Object1+ " ***** "+result2);
			}
			break;
				         
		case "Static":
		{
			wait.until(ExpectedConditions.alertIsPresent());
			String text=driver.switchTo().alert().getText();
			if(text.equalsIgnoreCase(ProValue1)|| text.equalsIgnoreCase(Parameter1))
			{
				ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"Dialog Message Exist",TestScenarioRow);
				System.out.println(Object1+ " ***** "+result1);
			}
			else
			{
				ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"Dialog Message does not Exist",TestScenarioRow);
				System.out.println(Object1+ " ***** "+result2);
			}
			driver.switchTo().alert().accept();
		}
		break;
		
		case "RTEditor":
            String text=driver.switchTo().frame(0).getPageSource();
            if(text.contains(Parameter1))
            {
            ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"FrameText Exist",TestScenarioRow);
            System.out.println(Object1+ " ***** "+result1);
            }
            else
            {
            ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"FrameText does not Exist",TestScenarioRow);
            System.out.println(Object1+ " ***** "+result2);
            }
            break;


		case "WebEdit":
			switch(Propname){
			case "name":
				xpath=driver.findElement(By.name(ProValue1));
				break;
			case "xpath":
				xpath=driver.findElement(By.xpath(ProValue1));
				break;
				}
			if(Parameter1==null)
			{
                                                if(xpath.isDisplayed())
                                              {
                                                     ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebEdit Exist",TestScenarioRow);
                                                     System.out.println(Object1+ " ***** "+result1);
                                              }
                                              else
                                              {
                                                     ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebEdit does not Exist",TestScenarioRow);
                                                     System.out.println(Object1+ " ***** "+result2);
                                              }
			}            
                
                  
                  else
                  {
                	  String WebEditval = xpath.getText();
                         String value=xpath.getAttribute("value");
                         String disabled=xpath.getAttribute("disabled");
	        String readonly=xpath.getAttribute("readonly");
                         if(Parameter1.equalsIgnoreCase(value)||Parameter1.equalsIgnoreCase(disabled)||value.contains(Parameter1)||Parameter1.equalsIgnoreCase(readonly)||Parameter1.equalsIgnoreCase(WebEditval))
                         {
                                ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebEdit Exist",TestScenarioRow);
                                System.out.println(Object1+ " ***** "+result1);
                         }
                         else
                         {
                                ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebEdit does not Exist",TestScenarioRow);
                                System.out.println(Object1+ " ***** "+result2);
                         }
                        
                  }
                 
       
            break;
			case "WebList":
			switch(Propname){
			case "name":
				xpath=driver.findElement(By.name(ProValue1));
				break;
				case "xpath":
				xpath=driver.findElement(By.xpath(ProValue1));
	
				break;

		
				}
		
			try{
			if(!Parameter1.equalsIgnoreCase(null))
			{
				String value=xpath.getAttribute("value");
				String disabled=xpath.getAttribute("disabled");
				String defaultvalue=xpath.getAttribute("value");
				String option=xpath.getText();
				Select visible=new Select(xpath);
				String visibletext=visible.getFirstSelectedOption().getText();
				if(Parameter1.equalsIgnoreCase(value)||Parameter1.equalsIgnoreCase(disabled)||Parameter1.equalsIgnoreCase(defaultvalue)||Parameter1.equalsIgnoreCase(option)||Parameter1.equalsIgnoreCase(visibletext)||value.contains(Parameter1)||option.contains(Parameter1)||(Parameter1.equals("Enable") && disabled==null))
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebList Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebList does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
				}
			}
			}
			catch(Exception e)
			{
				if(xpath.isDisplayed())
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,"WebList Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,"WebList does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
					}
			}
			break;
		case "WebCheckBox":	
		case "WebRadioGroup":
			switch(Propname){
			case "name":
				xpath=driver.findElement(By.name(ProValue1));
				break;
			case "xpath":
				xpath=driver.findElement(By.xpath(ProValue1));
				break;
			case "id":
				xpath=driver.findElement(By.id(ProValue1));
				break;
				}
			if(!(Parameter1==null))
			{
				String checked=xpath.getAttribute("checked");
				String disabled=xpath.getAttribute("disabled");
				String defaultvalue=xpath.getAttribute("value");
				Boolean selection=xpath.isSelected() ;
				if(Parameter1.equals(disabled)||Parameter1.equals(checked)||Parameter1.equals(selection)||Parameter1.equalsIgnoreCase(defaultvalue)||(Parameter1.equals("Enable") && disabled==null))
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,Object1 +"Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1 +"does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
				}
				}
				
			else
			{
				if(xpath.isDisplayed())
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,Object1 +"Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
				}
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1 +"does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
				}
			}
			break;
		case "Exception":
			String[] ProValue1arr;
			if(ProValue1.contains(","))
			{
				ProValue1arr = ProValue1.split(",");
				for(int i=0;i<ProValue1arr.length;i++)
					{
				
					if(driver.getPageSource().contains(ProValue1arr[i]))
					{
						ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,Object1+" " +ProValue1arr[i]+" Exist",TestScenarioRow);
						System.out.println(Object1+ " ***** "+result1);
					}
					
					else
					{
						ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1+" " +ProValue1arr[i]+" does not Exist",TestScenarioRow);
						System.out.println(Object1+ " ***** "+result2);
					}
					
					}
			}
			else
			{
				if(driver.getPageSource().contains(ProValue1)|| driver.getPageSource().contains(Parameter1)) 
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result1,Object1 +" " +ProValue1+" Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result1);
				}
				
				else
				{
					ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1 +" " +ProValue1+" does not Exist",TestScenarioRow);
					System.out.println(Object1+ " ***** "+result2);
				}
			}
				break;

		}	
		}
		catch(Exception e)
		{
			ResultReporting.printresultcheckpoint(filePath,workbook,step_Report,checkPointsh,module1,result2,Object1 +"does not Exist",TestScenarioRow);
			System.out.println(Object1+ " ***** "+result2);
		}
	}
	}
