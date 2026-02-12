package EM;



import java.time.Duration;

/*Description: Retrieve window and frame name from testcase sheet and Handles Windows and Frames functionality*/

import java.util.Iterator;
import java.util.List;
import java.util.Set;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Window_Frame_Handling {
	
	
	static String Parentframename,iframename,framename,frameinsideframesname;

	public static String framehandle(String framename,WebDriver driver,String Parameter1) throws InterruptedException
	{
		WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(10));
		JavascriptExecutor jsExecutor = (JavascriptExecutor)driver;
		try{
			
				driver.switchTo().defaultContent();
				switch(framename)
				{
				case "tab_comp1":
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("tab_comp");
             break;

				case "f_close_levels":
                    driver.switchTo().frame("f_view_levels");
                    driver.switchTo().frame("f_close_levels");
             break;
				case "main_comp1":
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("main_comp");
             break;
			case "OthSearchResults":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("OthSearchResults");
             break;
			case "querying":
				driver.switchTo().frame("querying");
				break;
	
			case "Chart_Create":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("Chart_Create");
             break;
			case "f_query_main":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_main");
             break;
			case "CARestrictCEHDataHDRFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("CARestrictCEHDataHDRFrame");
             break;
			case "SplChartGroupsMain":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("SplChartGroupsMain");
             break;
             
				case "f_queryCriteria":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_queryCriteria");
             break;
             
				case "f_query_search":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_search");
             break;
             
				case "f_query_search_result":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_search_result");
             break;
             
				case "DischargeCheckListTab1_frame1":
                    driver.switchTo().frame("DischargeCheckList_frame2");
                    driver.switchTo().frame("DischargeCheckListTab1_frame1");
             break;
             
				case "CARoleBasedAccessRightsforTaskListFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("CARoleBasedAccessRightsforTaskListFrame");
             break;
             
				case "f_drug":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_detail");
                    driver.switchTo().frame("f_ord_detail");
                    driver.switchTo().frame("f_drug_detail");
                    driver.switchTo().frame("f_drug");
             break;
				case "f_queryResult":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_queryResult");
             break;
             
				case "MedicationAdministrationRightsHeaderFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("MedicationAdministrationRightsHeaderFrame");
					break;
					
				case "MedicationAdministrationRightsListFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("MedicationAdministrationRightsListFrame");
					break;


             case "f_query_details":
               driver.switchTo().frame("content");    
                 driver.switchTo().frame("f_query_details");
               break;
                                  
             case "f_apptdetails":    
                   driver.switchTo().frame("f_apptdetails");
                                    break;
                                  
             case "OPCancelChkoutSearch":
                 driver.switchTo().frame("content");    
                 driver.switchTo().frame("OPCancelChkoutSearch");
                 break;
             case "OPCancelChkoutResult":
                 driver.switchTo().frame("content");    
                 driver.switchTo().frame("OPCancelChkoutResult");
                 break;
                                  
             case "f_compound_detail":
                                  driver.switchTo().frame("content");
                                  driver.switchTo().frame("workAreaFrame");
                                  driver.switchTo().frame("orderDetailFrame");
                                  driver.switchTo().frame("mainFrame");
                                  driver.switchTo().frame("DetailFrame");
                                  driver.switchTo().frame("criteriaPlaceOrderFrame");
                                  driver.switchTo().frame("f_prescription");
                                  driver.switchTo().frame("f_compound_detail");
                                  break;
                case "f_drugs":
                 driver.switchTo().frame("content");
                 driver.switchTo().frame("workAreaFrame");
                 driver.switchTo().frame("orderDetailFrame");
                 driver.switchTo().frame("mainFrame");
                 driver.switchTo().frame("DetailFrame");
                 driver.switchTo().frame("criteriaPlaceOrderFrame");
                 driver.switchTo().frame("f_prescription");
                 driver.switchTo().frame("f_ivdetails");
                 driver.switchTo().frame("f_iv_drug_details");
                 driver.switchTo().frame("f_drug_list");
                 driver.switchTo().frame("f_drugs");
                 break;
                 
                 
             case "f_prescription_form":
                 driver.switchTo().frame("content");
                 driver.switchTo().frame("workAreaFrame");
                 driver.switchTo().frame("orderDetailFrame");
                 driver.switchTo().frame("mainFrame");
                 driver.switchTo().frame("DetailFrame");
                 driver.switchTo().frame("criteriaPlaceOrderFrame");
                 driver.switchTo().frame("f_prescription");
                 driver.switchTo().frame("f_prescription_form");
                 break;
                                  
             case "f_compound_button":
                                  driver.switchTo().frame("content");
                                  driver.switchTo().frame("workAreaFrame");
                                  driver.switchTo().frame("orderDetailFrame");
                                  driver.switchTo().frame("mainFrame");
                                  driver.switchTo().frame("DetailFrame");
                                  driver.switchTo().frame("criteriaPlaceOrderFrame");
                                  driver.switchTo().frame("f_prescription");
                                  driver.switchTo().frame("f_compound_button");
                                  break;
                           case "f_preview_buttons3" :
                                  driver.switchTo().frame("f_preview_buttons"); 
                       break;
                                  
                           case "f_detail2":	
                            driver.switchTo().frame("content");
                            driver.switchTo().frame("f_query_add_mod");
                           driver.switchTo().frame("f_tab_detail");
                           driver.switchTo().frame("f_detail");
                           break; 
               case "placeOrderDetailFrame2":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("placeOrderDetailFrame");
					break;

                      case "CommonToolbar1":
					driver.switchTo().frame("CommonToolbar");
					break;
                               
                          case "f_query_perf_mand":                                  	                                                  
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_perf_mand");
					break;
					
				    
                    case "top":                                  	                                                  
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("top");
					break;
					
                    case "f_query_criteria_token_info":                                  	                                                  
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("f_query_criteria_token_info");
    					break;
    					
                    case "results":                                  	                                                  
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("bottom");
    					driver.switchTo().frame("results");
    					break;
					
                          case "f_query_header":                                  	                                                  
          					driver.switchTo().frame("content");
          					driver.switchTo().frame("f_query_add_mod");
          					driver.switchTo().frame("f_query_header");
          					break;
          					
                          case "f_query_header1":                                  	                                                  
            					driver.switchTo().frame("content");
            					driver.switchTo().frame("f_query_header");
            					break;
            					
                          case "f_query_detail":                                  	                                                  
            					driver.switchTo().frame("content");
            					driver.switchTo().frame("f_query_add_mod");
            					driver.switchTo().frame("f_query_detail");
            					break;
            					
                          case "query":                                  	                                                  
            					driver.switchTo().frame("content");
            					driver.switchTo().frame("f_query_add_mod");
            					driver.switchTo().frame("query");
            					break;
					
                    case "query_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("query_criteria_frame");
					driver.switchTo().frame("query_frame");
					break;
					
                    case "query_criteria_frame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("query_criteria_frame");
    				break;
                    case "query_result_frame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("query_result_frame");
    				break;
    				
                    case "autotracktoopd":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("autotracktoopd");
    				break;
    					
    				
                    case "patient_sub4":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("patient_sub2");
    					break;
    					
                    case "patient_sub6":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("patient_sub4");
    					break;
    					
    					
                    case "NurUtPractQueryFrame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("NurUtPractQueryFrame");
    					break;
    					
                    case "master_frame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("master_frame");
    					break;
    					
                    case "MasterFrame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("MasterFrame");
    					break;
    					
                    case "button_frame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("button_frame");
    					break;
    					
                    case "detail_frame1":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("detail_frame");
    					break;
    					
                    case "HeaderFrame1":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("HeaderFrame");
    					break;
    					
                    case "NurUtPractResultFrame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("NurUtPractResultFrame");
    					break;
    					
                    case "f_drug_interaction":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("f_drug_interaction");
    					break;
    					
                    case "MenuModify":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("MenuModify");
    					break;
    					
                    case "MenuItemAdd":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("MenuItemAdd");
    					break;
    					
                    case "MealCategoryAdd":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("MealCategoryAdd");
    					break;
                       					
                    case "f_query_add_mod_result":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod_main");
    					driver.switchTo().frame("f_query_add_mod_result");
    					break;
    					
                    case "f_query_add_mod_result1":
    				   driver.switchTo().frame("f_query_add_mod_result");
    					break;
    					
                    case "f_query_add_mod_main":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod_main");
    					driver.switchTo().frame("f_query_add_mod");
    					break;
    					
                    case "lookup_header" :
                                     driver.switchTo().frame("lookup_header");
                             break; 
                             
                            case "lookup_footer":
                                     driver.switchTo().frame("lookup_footer");
                             break;
				case "clinic_sub":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("clinic_sub");
                    break;
                    
				case "practitioner_sub":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("practitioner_sub");
                    break;
                    
				case "unitFrame1":
					driver.switchTo().frame("criteria_f0"); 
					driver.switchTo().frame("details");
                    driver.switchTo().frame("unitFrame");
                    break;
                           
				case "details_f2":
                    driver.switchTo().frame("criteria_f0"); 
                    driver.switchTo().frame("details");
                    driver.switchTo().frame("details_f1");
                    break;
                    
				case "details_f3":
					driver.switchTo().frame("content"); 
                    driver.switchTo().frame("workAreaFrame"); 
                    driver.switchTo().frame("details");
                    driver.switchTo().frame("details_f1");
                    break;
                    
                    
				case "performing_locn_top":
                    driver.switchTo().frame("content");    
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("performing_locn_top");
                    break;
                    
				case "performing_locn_bottom":
                    driver.switchTo().frame("content");    
                    driver.switchTo().frame("f_query_add_mod");
                     driver.switchTo().frame("performing_locn_bottom");
                    break;
				case "f_query_add_mod_header":
                    driver.switchTo().frame("content");    
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("f_query_add_mod_header");
                    break;
				case "f_query_add_mod_detail":
                    driver.switchTo().frame("content");    
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("f_query_add_mod_detail");
                    break;
				case "placeOrderDetailFrame1":
                    driver.switchTo().frame("criteriaPlaceOrderFrame");
                    driver.switchTo().frame("placeOrderDetailFrame");
                    break;
				case "sectionChartHeaderFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("sectionChartHeaderFrame");
                    break;
				case "chartComponentHeaderFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("chartComponentHeaderFrame");
                    break;
				case "sectionChartSearchFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("sectionChartSearchFrame");
                    break;
				case "sectionChartBottomFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("sectionChartBottomFrame");
                    break;
				case "chartComponentListFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("chartComponentListFrame");
                    break;
				case "AuthorizeOrdersBottomLeft":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("AuthoriseOrderBottom");
                    driver.switchTo().frame("AuthorizeOrdersBottomLeft");
                    break;
                    
				case "ClinicalEventHistoryCriteriaFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("ClinicalEventHistoryCriteriaFrame");
                    break;
                    
				case "RecordFrame1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ChkListDetailFrame");
					driver.switchTo().frame("RecordFrame");
					break;
					
				case "f_select":
					 driver.switchTo().frame("content");
					 driver.switchTo().frame("resultFrame");
					 driver.switchTo().frame("f_select");
					 break;
					 
				case "RTEditor1":
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("RecClinicalNotesFrame");
						driver.switchTo().frame("RecClinicalNotesSecDetailsEditorFrame");
						driver.switchTo().frame("RecClinicalNotesRTEditorFrame");
						driver.switchTo().frame("RTEditor0");
						break;
						
				case "OTNotesSearchCriteriaFrame":
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("OTNotesSearchCriteriaFrame");
						break;
				case "OTNotesSearchResultFrame":
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("OTNotesSearchResultFrame");
						break;

				case "ChkListRecordFrame1":
	                    driver.switchTo().frame("content");
	              		driver.switchTo().frame("f_query_add_mod");
	              		driver.switchTo().frame("ChkListRecordFrame");
	              		break;

				case "ChkListRecordFrame":
              		driver.switchTo().frame("f_query_add_mod");
              		driver.switchTo().frame("ChkListRecordFrame");
              		break;
				
				case "frameIssueReturnHeader":
					 driver.switchTo().frame("content");
					 driver.switchTo().frame("f_query_add_mod");
					 driver.switchTo().frame("frameIssueReturnHeader");
					 break;
					 
				case "f_term_code_set":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_term_code_set");
					break;
				case "f_term_add_modify":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_term_add_modify");
					break;
				case "f_term_result_header":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_term_result_header");
					break;
				case "f_term_code_result":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_term_code_result");
					break;
				case "OthsearchCriteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("OthsearchCriteria");
					break;
				
				case "problemsframe1":
					try{
				driver.switchTo().frame("content");
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame("problemsframe1");
                                                                                      break;
				}
					catch(Exception e){
						driver.switchTo().frame("problemsframe1");
                                                                                                                                break;                           
					}
				
				case "frmTest4":
				driver.switchTo().frame("content");
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame(1);
				driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryActiveProblems.jsp')]")));
				break;
				
				case "MyRefNotSeenCriteriaFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("MyRefNotSeenCriteriaFrame");
					break;
					
				case "wardretmedicationqueryframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_criteria");
					driver.switchTo().frame("wardretmedicationqueryframe");
					break;
					
				case "wardretmedicationdrugframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_criteria");
					driver.switchTo().frame("wardretmedicationdrugframe");
					break;
					
				case "wardretmedicationremarksframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_criteria");
					driver.switchTo().frame("wardretmedicationremarksframe");
					break;
					
				case "wardretmedicationactionframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_criteria");
					driver.switchTo().frame("wardretmedicationactionframe");
					break;
					
				case "wardretmedicationbuttonframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_criteria");
					driver.switchTo().frame("wardretmedicationbuttonframe");
					break;
					
				case "retmedicationqueryframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("retmedicationqueryframe");
					break;
					
				case "retmedicationbuttonframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("retmedicationbuttonframe");
					break;
					
				case "retmedicationdrugframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("retmedicationdrugframe");
					break;
					
				case "retmedicationremarksframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("retmedicationremarksframe");
					break;
					
				case "retmedicationactionframe":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("retmedicationactionframe");
					break;
					
				case "NewBornCriteriaFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("NewBornCriteriaFrame");
					break;
				
				case "NewBornResultFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("NewBornResultFrame");
					break;
					
				case "AssignRelationshipFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("AssignRelationshipFrame");
					break;
					
					
				case "detail_frame":
					driver.switchTo().frame(1);
					break; 
					
				case "criteriaFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("criteriaFrame");
					break;
					
				case "criteriaFrame1":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("criteriaFrame");
					break;
					
				case "resultFrame1":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("resultFrame");
					break;
					
				case "result4":
					driver.switchTo().frame("queries");	
					driver.switchTo().frame("result");
					break;
					
				case "result5":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("resultu");	
					driver.switchTo().frame("result");
					break;
					
				case "RecDiagnosisCurrentDiag1":	
					driver.switchTo().frame("RecDiagnosisMain");
					driver.switchTo().frame("RecDiagnosisCurrentDiag");
					break;
					
				case "MyRefNotSeenDetailsFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("MyRefNotSeenDetailsFrame");
					break;
				case "BallardScoreButtonFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("BallardScoreButtonFrame");
					break;
					
				 case "RecClinicalNotesTemplateFrame2":
						driver.switchTo().frame("RecClinicalNotesTemplateFrame");
						
						break; 
				 case "CASectionTemplatePreviewFrame":
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("CASectionTemplatePreviewFrame");
						break; 
						
						
				case "BallardScoreFrame":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("BallardScoreFrame");
					break;
					
				case "reg_prescriptions_footer_details2" :
	                 driver.switchTo().frame("content");
	                 driver.switchTo().frame("f_query_add_mod");
	                 driver.switchTo().frame("reg_prescriptions_header");
	                 driver.switchTo().frame("reg_prescriptions_footer_details");
	                 driver.switchTo().frame("reg_prescriptions_footer_details2");
	                 break;
	                 
				case "Docresult_add_mod":
					driver.switchTo().frame("content");
					driver.switchTo().frame("Docresult_add_mod");
					break;
				
				case "frameTaskForRespRelnHdr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameTaskForRespRelnHdr");
					break;
					
				case "frameTaskForRespRelnResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameTaskForRespRelnResult");
					break;
					
				case "frameTaskForRespRelnDtl":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameTaskForRespRelnDtl");
					break;
					
				case "menuFr":                                  	                                                  
						driver.switchTo().frame("menucontent");
						driver.switchTo().frame("menuFr");	
				break;
				
				case "patient_main":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("patient_main");
					break;
				case "Select_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("Select_frame");
					break;
				case "ChildBaseFrame":
					driver.switchTo().frame("RecordFrame");
					driver.switchTo().frame("ChildBaseFrame");
					break;
				case "ResultsFrame":
					driver.switchTo().frame("RecordFrame");
					driver.switchTo().frame("ResultsFrame");
					break;
					
				case "RecordFrame":
					driver.switchTo().frame("RecordFrame");
					driver.switchTo().frame("ChildBaseFrame");
					driver.switchTo().frame("RecordFrame");
					break;
					
				case "RecordFrame2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("RecordFrame");
					break;
					
				case "f_patientDiagnosisCriteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_patientDiagnosisCriteria");
					break;
					
				case "pend_ord_status_category":
					driver.switchTo().frame("content");
					driver.switchTo().frame("pend_ord_status_category");
					break;
					
				case "pend_ord_status_discharge":
					driver.switchTo().frame("content");
					driver.switchTo().frame("pend_ord_status_discharge");
					break;
					
				case "patient_main2":
					driver.switchTo().frame("detail");
					driver.switchTo().frame("patient_main");
					break;
					
				case "patient_main1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("patient_main");
				break;
				case "patient_main3":
					driver.switchTo().frame("patient_main");
				break;
				case "patient_main4":
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("patient_main");
				break;
					case "Frame61":
					driver.switchTo().frame("MainFrame2");
					driver.switchTo().frame("Frame61");
					break;
					
			case "messageFrame":
				try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("messageFrame");
					break;
				}
				catch(Exception e)
				{
					driver.switchTo().defaultContent();
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("messageFrame");
					break;
				}
				
			case "TriageButtonsFrame1":
				driver.switchTo().frame("content");
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame("TriageButtonsFrame");
				break;
	                 
				case "auditTrialFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("auditTrialFrame");
					break;
				case "FMGenPullListCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("FMGenPullListCriteriaFrame");
					break;
				case "RecDiagnosisCurrentDiag":	
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecDiagnosisMain");
					driver.switchTo().frame("RecDiagnosisCurrentDiag");
					break;

				case "main":
					driver.switchTo().frame("main");
					break;
					
				case "ModifyMealPlanCud":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ModifyMealPlanCud");
					break;
				case "ReleaseBed_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ReleaseBed_criteria");
					break;
				case "ModifyMealPlanSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					
					driver.switchTo().frame("ModifyMealPlanSearch");
				break;
				case "f_patientDiagnosisResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_patientDiagnosisResult");
					break;
				case "ModifyMealPlanItem":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
				
					driver.switchTo().frame("ModifyMealPlanItem");
					break;
				case "FMConfirmPullingListCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("FMConfirmPullingListCriteriaFrame");
					break;
				case "FMConfirmPullingListResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("FMConfirmPullingListResultFrame");
					break;
				case "FMGenPullListResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("FMGenPullListResultFrame");
					break;
				case "problemsframe0":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("problemsframe0");
					break;
					}
					catch(Exception e){
						driver.switchTo().frame("problemsframe0");
						break;
					}

					case "regOutRef_Criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("regOutRef_Criteria");
					break;
				case "Referral_Detail_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("Referral_Detail_frame");
					break;
				
				case "record_pendingorders_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("at_tab_frame");
					driver.switchTo().frame("record_pendingorders_frame");
					break;
				case "record_anethesia_details_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("at_tab_frame");
					driver.switchTo().frame("record_anethesia_details_frame");
					break;
					
				case "patient_sub":
							driver.switchTo().frame("content");
							driver.switchTo().frame("f_query_add_mod");
							driver.switchTo().frame("patient_sub");
							break;
							
				case "patient_sub1":
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("patient_sub");
						break;
					
				case "patient_sub2":
					driver.switchTo().frame("detail");
					driver.switchTo().frame("patient_sub");
					break;
			
				case "patient_sub5":
					driver.switchTo().frame("patient_sub");
					break;

				case "f_unsch_cases_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_tab_frames");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_unsch_cases_frame");
					break;
				case "f_swab_count_confirm_item":
					driver.switchTo().frame("f_tab_detail_frames");
					driver.switchTo().frame("f_swab_count_confirm_item");
					break;
				case "f_swab_count_dtls":
					driver.switchTo().frame("f_tab_detail_frames");
					driver.switchTo().frame("f_swab_count_dtls");
					break;
				case "f_swab_count_record":
					driver.switchTo().frame("f_tab_detail_frames");
					driver.switchTo().frame("f_swab_count_record");
					break;

					
	                case "Referral_Detail_frame1":
						driver.switchTo().frame("Referral_Detail_frame");
						break;
				
						
	                case "preview_editor":
	                	driver.switchTo().frame("content");
	                	driver.switchTo().frame("f_query_add_mod");
	                	driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("preview_editor");
						break;
	                case "OutstanListButton":
					driver.switchTo().frame("content");
					driver.switchTo().frame("issue_detail");
					driver.switchTo().frame("OutstanListButton");
					break;

				case "patient_sub3":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("patient_sub3");
					break;
                case "NPBOptionsFrame":
                	              driver.switchTo().frame("content");
                                       driver.switchTo().frame("workAreaFrame");
                                       driver.switchTo().frame("orderDetailFrame");
                                       driver.switchTo().frame("mainFrame");
                                       driver.switchTo().frame("DetailFrame");
                                       driver.switchTo().frame("criteriaPlaceOrderFrame");
                                       driver.switchTo().frame("f_prescription");
                                       driver.switchTo().frame("NPBOptionsFrame");
                                       break;
                                       
                            
                                       
                              case "f_query_add_mod5":
                            	  try{
                                       driver.switchTo().frame("content");
                                       driver.switchTo().frame("workAreaFrame");
                                       driver.switchTo().frame("orderDetailFrame");
                                       driver.switchTo().frame("mainFrame");
                                       driver.switchTo().frame("DetailFrame");
                                       driver.switchTo().frame("criteriaPlaceOrderFrame");
                                       driver.switchTo().frame("f_query_add_mod");
                                       break;
                            	  }
                            	  catch(Exception e)
                            	  {
                            		  try{
                            			  driver.switchTo().defaultContent();
                            			  driver.switchTo().frame("content");
                                          driver.switchTo().frame("workAreaFrame");
                                          driver.switchTo().frame("orderDetailFrame");
                                          driver.switchTo().frame("mainFrame");
                                          driver.switchTo().frame("DetailFrame");
                                          driver.switchTo().frame("criteriaPlaceOrderFrame");
                                          driver.switchTo().frame("f_prescription");
                                          driver.switchTo().frame("f_query_add_mod");
                                          break;
                            		  }
                            		  catch(Exception e1)
                                	  {
                            		  driver.switchTo().defaultContent();
                            		  driver.switchTo().frame("f_query_add_mod");
                                      break;
                                	  }
                            	  }
                                
                              case "editor_button1":
					driver.switchTo().frame("editor_button");
					break;
					
				case "frame1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("frame1");
					driver.switchTo().frame("frame1");
					break;
				case "frame5":
					driver.switchTo().frame("content");
					driver.switchTo().frame("frame1");
					break;
				case "MenuItemPopulate":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("MenuItemPopulate");
					break;
					
				case "frame2":
					try{
						driver.switchTo().frame("content");
						driver.switchTo().frame("frame2");
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();	
						driver.switchTo().frame("content");
						driver.switchTo().frame("frame1");
						driver.switchTo().frame("frame2");
					}
					break;
				case "Main_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("Main_frame");
					break;
                  
				case "Main_frame1":
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("Main_frame");
					break;
				case "searchbutton":
					driver.switchTo().frame("content");
					driver.switchTo().frame("frame1");
					driver.switchTo().frame("searchbutton");
					break;
				case "Pending Orders":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryPendingOrdersD.jsp')]")));
					break;
									
				case "PatToolbarFr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("PatToolbarFr");
					break;
				
				case "PatSearchCriteriaFr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("PatSearchCriteriaFr");
					break;
					
				case "commontoolbarFrame":
					try
					{
					driver.switchTo().frame("content");
					driver.switchTo().frame("commontoolbarFrame");
					}
					catch(Exception e)
					{
						driver.switchTo().frame("commontoolbarFrame");
					}
				break;
				case "commontoolbarFrame1":
					
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("commontoolbarFrame");
					
				break;
				
				case "commontoolbarFrame2":
					
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("commontoolbarFrame");
					break;
					
					
				case "commontoolbarFrame3":
					driver.switchTo().frame("commontoolbarFrame");
					break;
					
				case "RecClinicalNotesSearchDetailsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesSearchDetailsFrame");
				break;
				
				case "RecClinicalNotesSearchDetailsFrame1":
					driver.switchTo().frame("RecClinicalNotesSearchDetailsFrame");
				break;
			
				
				case "RecClinicalNotesSearchToolbarFrame1":
					driver.switchTo().frame("RecClinicalNotesSearchToolbarFrame");
				break;
				
				case "ReviewNotesCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ReviewNotesCriteriaFrame");
				break;	
				
				case "ReviewNotesDetailsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ReviewNotesDetailsFrame");
				break;	
				
				case "content":
					driver.switchTo().frame("content");
					driver.switchTo().frame("content");
					break;
					
				case "content1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("resultFrame");
					driver.switchTo().frame("content");
					break;
			
				case "lookup_head":   
					driver.switchTo().frame("lookup_head");
					break;
				case "lookup_detail":
					driver.switchTo().frame("lookup_detail");
					break;
					
				case "blank":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("blank");
					break;
					
				case "f_query_add_mod2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_add_mod");
					break;
					
				case "f_query_add_mod3":
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("f_query_add_mod");
						break;
						
				case "f_query_add_mod_query":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("	");
					break;
					
				case "f_query_add_mod_res":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_add_mod_res");
					break;	
					
				case "f_query_add_mod0":	
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame(0);
						driver.switchTo().frame("f_query_add_mod_res");
						break;	
						
				case "f_query_add_mod1":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod1");
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("f_query_add_mod1");
					}
					break;
				
					
				case "f_query_add_mod":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					break;
					
				case "f_query_clntaudittrialresult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_clntaudittrialresult");
					break;
					
				case "f_query_add_mod6":
					driver.switchTo().frame("f_query_add_mod");
					break;
				
				case "f_query_add_mod_display":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("f_query_add_mod_display");
					break;
					
				case "f_query_add_mod_display1":
					driver.switchTo().frame("f_query_add_mod_display");
					break;
					
				case "f_query_add_mod4":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("f_query_add_mod");
					break;
				
				case "f_query_add_mod_button1":
					driver.switchTo().frame("f_query_add_mod_button");
					break;
					
				case "f_query_add_mod_cancel":
					driver.switchTo().frame("content");
					driver.switchTo().frame(1);
					break;
					
				case "f_query_add_mod_EncID":
					driver.switchTo().frame("content");
					driver.switchTo().frame(driver.findElement(By.xpath("//frame[contains(@src,'../../eCommon/html/blank.html')]")));
					break;
					
				case "OPCheckoutDeath":
					driver.switchTo().frame(driver.findElement(By.xpath("//frame[contains(@src,'../../eCommon/html/blank.html')]")));
					break;				
					
				case "search_static":
					driver.switchTo().frame("content");
					driver.switchTo().frame("search_static");
					break;
					
				case "TaskListleftStatusFrame" :
		          	   driver.switchTo().frame("content");
		          	   driver.switchTo().frame("workAreaFrame");
		          	   driver.switchTo().frame("TaskListLeftFrame");
		               driver.switchTo().frame("TaskListleftStatusFrame");
		                 break;
			
					
				case "ResultEntryBtn":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ResultEntryBtn");
					break;
				
					
				case "PatCriteriaFr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("PatCriteriaFr");
					break;
					
				case "heading":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ViewClinicalNoteNoteContentMainDetailFrame");
					driver.switchTo().frame("heading");
					break;
					
				case "ChartRecordingDetailFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ChartRecordingDetailFrame");
					break;
					
				case "ChartRecordingDetailFrame1":
					driver.switchTo().frame("ChartRecordingDetailFrame");
					break;
					
				case "ChartRecordingToolBarFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ChartRecordingToolBarFrame");
					break;
					
				case "ChartRecordingToolBarFrame1":
					driver.switchTo().frame("ChartRecordingToolBarFrame");
					break;
					
				case "RTEditor0":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					driver.switchTo().frame("RTEditor0");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("RecClinicalNotesFrame");
						driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
						driver.switchTo().frame(0);
						break;
					}
					
							
				case "RecClinicalNotesToolbarFrame":
				  	driver.switchTo().frame("content");
				  	driver.switchTo().frame("workAreaFrame");
				  	driver.switchTo().frame("RecClinicalNotesFrame");
				  	driver.switchTo().frame("RecClinicalNotesToolbarFrame");
				  	break;	
				
				  	
				case "ChartRecordingControlsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ChartRecordingControlsFrame");
					break;
					
				case "controlsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("details");
					driver.switchTo().frame("controlsFrame");
					break;
					
				case "ChartRecordingCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ChartRecordingCriteriaFrame");
					break;
					
				case "ChartRecordingCriteriaFrame1":
					driver.switchTo().frame("ChartRecordingCriteriaFrame");
					break;
					
				case "CA-InpatRef":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("f_query_add_mod");
					break;
					
				case "code_label":
					driver.switchTo().frame("TermCodeSearchByFrame");
					driver.switchTo().frame("code_label");
					break;
					
				case "codedesc":
					driver.switchTo().frame("TermCodeSearchByFrame");
					driver.switchTo().frame("codedesc");
					break;
					
				case "CusticdQueryResultFrame":
					driver.switchTo().frame("TermCodeSearchByFrame");
					driver.switchTo().frame("code_description");
					driver.switchTo().frame("CusticdQueryResultFrame");
					break;
					
				case "LocnResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("PatResultFr");
					driver.switchTo().frame("LocnResult");
					break;
				case "LocnResultPatClass":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("PatResultFr");
					driver.switchTo().frame("LocnResultPatClass");
					break;
					
				case "RecPatChiefComplaintAddModifyFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecPatChiefComplaintAddModifyFrame");
					break; 
					
				case "RecPatChiefComplaintResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecPatChiefComplaintResultFrame");
					break; 
					
				case "SpecimenOrderSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("SpecimenOrderSearch");
					break; 
					
				case "SpecimenOrderResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("SpecimenOrderResult");
					break; 
				
				case "TaskListRightResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("TaskListRightFrame");
					driver.switchTo().frame("TaskListRightResultFrame");
					break;
					
					case "Transfer_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("Transfer_frame");
					break;
					
					case "Transfer_frame1":
					driver.switchTo().frame("Transfer_frame");
					break;
				
				case "Transfer_in":
					driver.switchTo().frame("Transfer_in");
					break;
					
			
					
				case "CommonToolbar":
					driver.switchTo().frame("content");
					driver.switchTo().frame("CommonToolbar");
					break;
					
				case "menuFrame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("menuFrame");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("menuFrame");
						break;
					
					}
				case "orderMainTab":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderMainTab");
					break;
					
				case "orderMainTab2":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderMainTab");
					break;
					
				case "criteriaDetailFrame1":
					try{
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					break;
					}
	
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("criteriaPlaceOrderFrame");
						driver.switchTo().frame("criteriaDetailFrame");
						break;
					}
					
			
					
				case "resultDtlFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					driver.switchTo().frame("resultDtlFrame");
					break;
					
				case "orderMainTab1":
					driver.switchTo().frame("orderMainTab");
					break;
					
				case "issue_header":
					driver.switchTo().frame("content");
					driver.switchTo().frame("issue_header");
					break;
					
				case "apptdairy":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("apptdairy");
					break;
					
				case "apptdairy1":
					driver.switchTo().frame("apptdairy");
					break;
					
				case "tabFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("tabFrame");
					break;
					
				case "issue_tab":
					driver.switchTo().frame("content");
					driver.switchTo().frame("issue_tab");
					break;
					
				case "CAHealthRiskAssessmentDiagnosisFrm":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("CAHealthRiskAssessmentDiagnosisFrm");
					break;
					
				case "CAHealthRiskAssessmentOrderCatalogFrm":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("CAHealthRiskAssessmentOrderCatalogFrm");
					break;
					
				case "ResultReportingSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ResultReportingSearch");
					break;
					
				case "frameMultiPatientOrdersHdrResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameMultiPatientOrdersHdrResult");
					break;
					
				case "ExistingOrderSearch2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderMainTab");
					driver.switchTo().frame("ExistingOrderSearch");
					break;
					
				case "messageFrame1":
					try
					{
					driver.switchTo().frame("messageFrame");
					break;
					}
					catch(Exception e)
					{
					driver.switchTo().defaultContent();
					driver.switchTo().frame("content");
		            driver.switchTo().frame("messageFrame"); 
		                 break;
					}  
		                 
		                 
				case "messageFrame2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("messageFrame");
					break;
					
				case "ResultReportingResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ResultReportingResult");
					driver.switchTo().frame("ResultReportingResult");
					break;
					
				case "RecClinicalNotesRTEditorFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsEditorFrame");
					driver.switchTo().frame("RecClinicalNotesRTEditorFrame");
					break;
					
				case "externalOrders":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameExternalOrders");
					driver.switchTo().frame("externalOrders");
					break;	
					
				case "ExistingOrderSearch1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameExternalOrders");
					driver.switchTo().frame("externalOrdersResult");
					driver.switchTo().frame("ExistingOrderSearch");
					break;
					
				case "ExistingOrderResult1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameExternalOrders");
					driver.switchTo().frame("externalOrdersResult");
					driver.switchTo().frame("ExistingOrderResult");
					driver.switchTo().frame("ExistingOrderResult");
					break;	
					
				case "criteriaMainFrame1":	
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("criteriaMainFrame");
					break;
					
				case "motherline_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("motherline_frame");
					break;
					
				case "criteriaMainFrame2":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaMainFrame");
					break;
					
					
				case "criteriaMainFrame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaMainFrame");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("orderDetailFrame");
						driver.switchTo().frame("mainFrame");
						driver.switchTo().frame("DetailFrame");
						driver.switchTo().frame("criteriaPlaceOrderFrame");
						driver.switchTo().frame("criteriaMainFrame");
						break;
					}
					
				case "criteriaMainFrame3":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaMainFrame");
					break;
					
				case "oncology_options":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("oncology_options");
					break;
		                 
				case "oncology_button":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("oncology_drugs");
					driver.switchTo().frame("oncology_button");
					break;
					
				case "oncology_button1":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("oncology_detail");
					driver.switchTo().frame("oncology_drugs");
					driver.switchTo().frame("oncology_button");
					break;
					
				case "f_preview_buttons" :
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
		           	driver.switchTo().frame("criteriaPlaceOrderFrame");
		           	driver.switchTo().frame("f_prescription");
		           	driver.switchTo().frame("oncology_drugs");
		            driver.switchTo().frame("f_preview_buttons"); 
		         break;
		         
				case "f_preview_buttons1" :
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
		           	driver.switchTo().frame("criteriaPlaceOrderFrame");
		           	driver.switchTo().frame("oncology_detail");
		           	driver.switchTo().frame("oncology_drugs");
		            driver.switchTo().frame("f_preview_buttons"); 
		         break;
		         
				case "oncology_admin" :
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
		           	driver.switchTo().frame("criteriaPlaceOrderFrame");
		           	driver.switchTo().frame("f_prescription");
		           	driver.switchTo().frame("oncology_drugs");
		            driver.switchTo().frame("oncology_admin"); 
		            break;
		            
				case "oncology_admin1" :
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
		           	driver.switchTo().frame("criteriaPlaceOrderFrame");
		           	driver.switchTo().frame("oncology_detail");
		           	driver.switchTo().frame("oncology_drugs");
		            driver.switchTo().frame("oncology_admin"); 
		            break;
		      
				case "resultHdrFrame":
					try{
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					driver.switchTo().frame("resultHdrFrame");
					break;
					}
				catch(Exception e)
					{
					driver.switchTo().defaultContent();
					 driver.switchTo().frame("content");
					 driver.switchTo().frame("workAreaFrame");
					 driver.switchTo().frame("orderDetailFrame");
					 driver.switchTo().frame("mainFrame");
					 driver.switchTo().frame("DetailFrame");
					 driver.switchTo().frame("criteriaDetailFrame");
					 driver.switchTo().frame("resultHdrFrame");
		                 break;
					}
					
				case "resultDtlFrame1":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("resultDtlFrame");
					break;
					
				case "resultDtlFrame2":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					driver.switchTo().frame("resultDtlFrame");
					break;
					
				case "resultDtlFrame3":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					driver.switchTo().frame("resultDtlFrame");
					break;
					
				case "resultDtlFrame4":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaTickSheetsFrame");
					driver.switchTo().frame("resultDtlFrame");
					break;
					
				case "resultListFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					driver.switchTo().frame("resultListFrame");
					break;
					
				case "tabFrame1":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("tabFrame");
					break;
					
				case "tabFrame2":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("tabFrame");
					break;
					
					case "frame4":
					driver.switchTo().frame(1);
					break;
					
					case "serachFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("serachFrame");
					break;
					
				case "adv_bed_mgmt":
					driver.switchTo().frame("content");
					driver.switchTo().frame("adv_bed_mgmt");
					break;
					
				case "ExistingOrderSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("ExistingOrderSearch");
					break;

				case "ExistingOrderSearch3":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("ExistingOrderSearch");
					break;
					
				case "ConsentOrderTop":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ConsentOrderTop");
					break;
					
				case "ConsentOrderTop1":
					driver.switchTo().frame("ConsentOrderTop");
					break;
					
				case "ExistingOrderResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("ExistingOrderResult");
					driver.switchTo().frame("ExistingOrderResult");
					break;
					
				case "ExistingOrderResult3":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("ExistingOrderResult");
					driver.switchTo().frame("ExistingOrderResult");
					break;
					
				case "ExistingOrderResult2":
					try
					{
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("ExistingOrderResult");
					driver.switchTo().frame("ExistingOrderResult");
					break;
				}
				catch(Exception e)
				{
					driver.switchTo().defaultContent();
					driver.switchTo().frame("ExistingOrderResult");
					driver.switchTo().frame("ExistingOrderResult");
	                 break; 
				}   
				
				case "ConsentOrdersBottomRight1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ConsentOrderBottom");
					driver.switchTo().frame("ConsentOrdersBottomRight");
					driver.switchTo().frame("ConsentOrdersBottomRight1");
					break;
					
				case "ConsentOrdersBottomRight1_1":
					driver.switchTo().frame("ConsentOrderBottom");
					driver.switchTo().frame("ConsentOrdersBottomRight");
					driver.switchTo().frame("ConsentOrdersBottomRight1");
					break;
					
				case "ConsentOrdersBottomRight2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ConsentOrderBottom");
					driver.switchTo().frame("ConsentOrdersBottomRight");
					driver.switchTo().frame("ConsentOrdersBottomRight2");
					break;
					
				case "ConsentOrdersBottomRight2_1":
					driver.switchTo().frame("ConsentOrderBottom");
					driver.switchTo().frame("ConsentOrdersBottomRight");
					driver.switchTo().frame("ConsentOrdersBottomRight2");
					break;
					
				case "ConsentOrdersBottomRight3":
					driver.switchTo().frame("ConsentOrderBottom");
					driver.switchTo().frame("ConsentOrdersBottomRight");
					driver.switchTo().frame("ConsentOrdersBottomRight2");
					break;	
					
				case "ConsentOrdersBottomRight4":
					driver.switchTo().frame("ConsentOrderBottom");
					driver.switchTo().frame("ConsentOrdersBottomRight");
					driver.switchTo().frame("ConsentOrdersBottomRight1");
					break;
					
				case "AuthorizeOrdersBottomRight1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("AuthoriseOrderBottom");
					driver.switchTo().frame("AuthorizeOrdersBottomRight");
					driver.switchTo().frame("AuthorizeOrdersBottomRight1");
					break;
					
				case "AuthoriseOrderTop":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("AuthoriseOrderTop");
					break;
					
				case "RegisterOrderSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RegisterOrderSearch");
					break;
					
				case "RegisterOrderResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RegisterOrderResult");
					break;
					
				case "ViewOrderHistoryResult":
					driver.switchTo().frame("ViewOrderHistoryResult");
					break;
					
				case "RegisterOrderBtn":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RegisterOrderBtn");
					break;
					
				case "frame_Procedures":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frame2");
					break;
			
					
				case "frame3":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frame3");
					break;
					
				case "frame1_Procedures":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frame1");
					break;
					
				case "AuthorizeOrdersBottomRight2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("AuthoriseOrderBottom");
					driver.switchTo().frame("AuthorizeOrdersBottomRight");
					driver.switchTo().frame("AuthorizeOrdersBottomRight2");
					break;
					
				case "f_query_result1":
					try{
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("f_query_result1");
					break;
					}
					catch(Exception e){
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("f_query_result");
						driver.switchTo().frame("f_query_result1");
						break;
					}
				case "f_query_result1_1":
					driver.switchTo().frame("med_reconc_results");
					driver.switchTo().frame("f_query_result1");
					break;
					
				case "f_query_result1_2":
					driver.switchTo().frame("med_reconc_results");
					driver.switchTo().frame("f_query_result");
					break;
					
				case "f_query_result":
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("f_query_result");
					break;
					
				case "f_query_result2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_result");
					break;
					
				case "Docquery_add_mod":
					driver.switchTo().frame("content");
					driver.switchTo().frame("Docquery_add_mod");
					break;
					
			case "f_detail":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_detail");
					break;
					
			case "f_detail3":
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame("orderDetailFrame");
				driver.switchTo().frame("mainFrame");
				driver.switchTo().frame("DetailFrame");
				driver.switchTo().frame("criteriaPlaceOrderFrame");
				driver.switchTo().frame("f_detail");
				break;
					
				case "f_detail1":
					try
					{
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_detail");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
		          	   driver.switchTo().frame("content");
		          	 driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("f_tab_detail");
		               driver.switchTo().frame("f_detail");
		                 break; 
					}   
					
				case "f_button1":
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("orderDetailFrame"));
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_button");
					break;
					
					case "CAConfidentialitySetUpHeaderFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("CAConfidentialitySetUpHeaderFrame");
					break;
					
					case "CAConfidentialitySearchQueryFrame":
						driver.switchTo().frame("content");
						driver.switchTo().frame("CAConfidentialitySearchQueryFrame");
						break;
				
					
				case "CAConfidentialitySearchResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("CAConfidentialitySearchResultFrame");
					break;
					
				case "placeOrderDetailFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("placeOrderDetailFrame");
					break;
					
				case "placeOrderDetailFrame3":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("placeOrderDetailFrame");
					break;
					
				case "RecClinicalNotesTabsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesTabsFrame");
					break;
					
						
				case "RecClinicalNotesSrchResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSrchResultFrame");
					break; 
					
				case "RecClinicalNotesSrchCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSrchCriteriaFrame");
					break;
					
				case "RecDiagnosisAddModify1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecDiagnosisMain");
					driver.switchTo().frame("RecDiagnosisAddModify");
					break; 
					
				case "RecDiagnosisOpernToolbar1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecDiagnosisMain");
					driver.switchTo().frame("RecDiagnosisOpernToolbar");
					break;
					
				case "DisplayCriteria":
					driver.switchTo().frame("criteria_f0");
					driver.switchTo().frame("details");
					driver.switchTo().frame("details_f2");
					driver.switchTo().frame("DisplayResult");
					driver.switchTo().frame(1);
					break;
					
				case "DisplayCriteria1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("details");
					driver.switchTo().frame("details_f2");
					driver.switchTo().frame("DisplayResult");
					driver.switchTo().frame(1);
					break;
					
				case "dataFrame1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("details");
					driver.switchTo().frame("dataFrame");
					break;
					
		
					
				case "encountercontrol":
					driver.switchTo().frame("criteria_f0");
					driver.switchTo().frame("details");
					driver.switchTo().frame("details_f2");
					driver.switchTo().frame("encountercontrol");
				break;
				
				case "encountercontrol_1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("details");
					driver.switchTo().frame("details_f2");
					driver.switchTo().frame("encountercontrol");
				break;
				
				case "encountercontrol_2":
					driver.switchTo().frame("details");
					driver.switchTo().frame("details_f2");
					driver.switchTo().frame("encountercontrol");
				break;
				
				case "dataFrame":
					driver.switchTo().frame("criteria_f0");
					driver.switchTo().frame("details");
					driver.switchTo().frame("dataFrame");
				break;
			
			
				case "code_label1":
					driver.switchTo().frame("group_blank");
					driver.switchTo().frame("code_label");
					break; 
					
				case "codedesc1":
					driver.switchTo().frame("group_blank");
					driver.switchTo().frame("codedesc");
					break;
					
				case "criteriaDetailFrame2":
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					break;
					
				case "Transfer_frame2":
					driver.switchTo().frame(2);
					break;
					
				case "criteriaDetailFrame3":	
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("criteriaDetailFrame");
						break;
					}
				case "criteriaDetailFrame4":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					break;
			
				case "criteriaDetailFrame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaDetailFrame");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("orderDetailFrame");
						driver.switchTo().frame("mainFrame");
						driver.switchTo().frame("DetailFrame");
						driver.switchTo().frame("criteriaPlaceOrderFrame");
						driver.switchTo().frame("criteriaDetailFrame");
						break;
					}
					
				case "result":
					try{
						driver.switchTo().frame("result");
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("queries");
						driver.switchTo().frame("result");
					}
						break;
						
	case "result_1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("result");
					break;
					
				case "result2":
					try{
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("ViewClinicalNoteNoteContentMainDetailFrame");
						driver.switchTo().frame("result");
					
						break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("result");
						break;
					}
				      
				
					
					
				case "f_query_criteria":
                    try{
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_criteria");
                    }
                    catch(Exception e)
                    {
                           try{
                                  driver.switchTo().frame("f_query_add_mod");
                                  driver.switchTo().frame("f_query_criteria");
                           }
                           catch(Exception e1)
                           {
                                  driver.switchTo().defaultContent();
                                  driver.switchTo().frame("f_query_criteria");
                           }
                    }
                    break;

				case "addmodify":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("addmodify");
					break;
					
				case "AddModify":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("AddModify");
					break;
					
				case "noteTypeValuesFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("noteTypeValuesFrame");
					break;
					
				case "addSectionsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("addSectionsFrame");
					break;
					
				case "ChiefComplaintMaster":
					driver.switchTo().frame("content");
					driver.switchTo().frame("ChiefComplaintMaster");
					break;
					
				case "ChiefComplaintDiag":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ChiefComplaintDiag");
					break;
					
				case "ChiefComplaintDiagSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ChiefComplaintDiagSearch");
					break;
					
				case "ChiefComplaintDiagDetails":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ChiefComplaintDiagDetails");
					break;
					
				case "drugMasterMain":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("drugMasterMain");
					break;
					
				case "f_add_modify":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_tab_detail");
					driver.switchTo().frame("f_add_modify");
					break;
					
					
				case "clinic_main":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("clinic_main");
					break;
					
				case "createvolume_header":
					driver.switchTo().frame("content");
					driver.switchTo().frame("createvolume_header");
					break;
					
					case "frameNoteTypeRespHdr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameNoteTypeRespHdr");
					break;
					
				case "frameNoteTypeRespDtl":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameNoteTypeRespDtl");
					break;
					
				case "frameNoteTypeRespResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameNoteTypeRespResult");
					break;
				
					
				case "result3":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("dummy");
					driver.switchTo().frame("result");
					break;
					}
					catch(Exception e)
                    {
						driver.switchTo().defaultContent();
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("result");
					break;
                    }
					
				case "patientFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("patientFrame");
					break;
					
				case "patientDetailsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("patientDetailsFrame");
					break;
					
				case "searchCriteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("searchCriteria");
					break;
					
				case "requester_details":
					driver.switchTo().frame("content");
					driver.switchTo().frame("requester_details");
					break;
					
				case "search_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("search_criteria");
					break;
					
				case "receipt_header":
					driver.switchTo().frame("content").switchTo().frame("receipt_header");
					break;
					
				case "cancel_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("cancel_criteria");
					break;
					
				case "cancel_details":
					driver.switchTo().frame("content");
					driver.switchTo().frame("cancel_details");
					break;
					
				case "summary":
					driver.switchTo().frame("content");
					driver.switchTo().frame("summary");
					break;	
					
				case "receipt_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("receipt_criteria");
					break;
					
				case "detail":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("detail");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().frame("detail");
						break;
					}
					
					
				case "detail2":
						driver.switchTo().frame("detail");
						break;
				
				case "detail1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("detail");
					break;
					
		
				case "querying1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("querying");
					break;
					

					
				case "Search":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("Search");
					break;
					
					
				case "ReviewResultsSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ReviewResultsSearch");
					break;
					
				case "ReviewResultsResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ReviewResultsResult");
					break;
					
				case "ReviewResultsBtn":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
				    driver.switchTo().frame("ReviewResultsBtn");
					break;
							
				case "result1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("result");
					Thread.sleep(1000);
					driver.switchTo().frame("result1");
					break;
					
				case "result12":
					driver.switchTo().frame("result");
					driver.switchTo().frame("result1");
					break;
				
					
				case "add_modify":
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("add_modify");
					break;
					
				case "dispFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("dispFrame");
					break;
					
				case "f_RX":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");					
					driver.switchTo().frame("orderDetailFrame");				
					driver.switchTo().frame("mainFrame");					
					driver.switchTo().frame("DetailFrame");					
					driver.switchTo().frame("criteriaPlaceOrderFrame");					
					driver.switchTo().frame("f_prescription");				
					driver.switchTo().frame(2);
					break;
					
				case "f_RX1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");					
					driver.switchTo().frame("orderDetailFrame");				
					driver.switchTo().frame("mainFrame");					
					driver.switchTo().frame("DetailFrame");					
					driver.switchTo().frame("criteriaPlaceOrderFrame");					
					driver.switchTo().frame("f_prescription");				
					driver.switchTo().frame("f_RX");
					break;
					
				case "f_RX2":
					driver.switchTo().frame("workAreaFrame");					
					driver.switchTo().frame("orderDetailFrame");				
					driver.switchTo().frame("mainFrame");					
					driver.switchTo().frame("DetailFrame");					
					driver.switchTo().frame("criteriaPlaceOrderFrame");					
					driver.switchTo().frame("f_prescription");				
					driver.switchTo().frame(2);
					break;
					
				case "f_ivbutton":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivbutton");
					break;
					
				case "f_header":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_header");
					break;
					
				case "f_ivselect":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivselect");
					break;
					
				case "f_ivdetails":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					break;
					
				case "f_sub_ivdrugs":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_drug_details");
					driver.switchTo().frame("f_sub_ivdrugs");
					break;
					
				case "f_preview":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_preview");
					break;
				
					
				case "criteriaButtonFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaButtonFrame");
					break;
					
				case "criteriaButtonFrame1":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaButtonFrame");
					break;
					
				case "criteriaPlaceOrderFrame":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_button");
					break;	
					
				case "f_drug_button":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_drug_details");
					driver.switchTo().frame("f_drug_button");
					break;
					
				case "f_oncology_regimen_drug_list":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("oncology_drugs");
					driver.switchTo().frame("f_oncology_regimen_drug_list");
					break;
					
				case "f_drug_button1":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_detail");
					driver.switchTo().frame("f_iv_drug_details");
					driver.switchTo().frame("f_drug_button");
					break;
					
				case "f_iv_fluid":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_fluid");
					break;
					
				case "f_iv_fluid1":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_detail");
					driver.switchTo().frame("f_iv_fluid");
					break;
					
				case "f_iv_admin":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_admin");
					break;
					
				case "f_iv_pb_drug":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_pb_drug");
					break;
					}
					catch(Exception e)
					{
					driver.switchTo().defaultContent();
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_drug_details");
					driver.switchTo().frame("f_sub_ivdrugs");	
					break;
					}
					
				case "f_iv_pb":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_pb");
					break;
					
				case "f_iv_pb_admin_dtls":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_admin");
					driver.switchTo().frame("f_iv_pb_admin_dtls");
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_prescription");
					driver.switchTo().frame("f_ivdetails");
					driver.switchTo().frame("f_iv_pb_admin_dtls");
					break;
					}
					
				case "f_iv_pb_admin_dtls1":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_detail");
					driver.switchTo().frame("f_iv_pb_admin_dtls");
					break;
					
				case "f_button":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_button");
					break;
					}
					catch(Exception e){
						driver.switchTo().frame("f_button");
						break;
					}
				case "f_button2":
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_button");
					break;
					
				case "f_button3":
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("criteriaPlaceOrderFrame");
					driver.switchTo().frame("f_button");
					break;
					
				case "patient_id_locator":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatFrame");
					driver.switchTo().frame("patient_id_locator");
					break;
					
				case "patient_details":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatDetFrame_1");
					driver.switchTo().frame("patient_details");
					break;
					
				case "order_details":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatDetFrame_1");
					driver.switchTo().frame("order_details");
					break;
					
				case "f_disp_medication_filling":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatDetFrame_3");
					driver.switchTo().frame("f_disp_medication_filling");
					break;
					
				case "f_disp_medication_verification":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatDetFrame_3");
					driver.switchTo().frame("f_disp_medication_verification");
					break;
					
				case "f_disp_medication_allocation":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatDetFrame_3");
					driver.switchTo().frame("f_disp_medication_allocation");
					break;
					
				case "f_disp_medication_all_stages":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_result");
					driver.switchTo().frame("DispMedicationPatDetFrame_3");
					driver.switchTo().frame("f_disp_medication_all_stages");
					break;
					
				case "CASectionTemplateHeaderFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("CASectionTemplateHeaderFrame");
					break;
				case "CASectionTemplateListFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("CASectionTemplateListFrame");
					break;
				case "CASectionTemplateDetailFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("CASectionTemplateDetailFrame");
					break;
				case "CASectionTemplateToolbarFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("CASectionTemplateDetailFrame");
					driver.switchTo().frame("CASectionTemplateToolbarFrame");
					break;
				case "QueryTemplateDetail":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("QueryTemplateDetail");
					break;
				case "criteria_f1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("criteria_f1");
					break;
					
				case "criteria_f1_2":
					driver.switchTo().frame("criteria_f0");
					driver.switchTo().frame("criteria_f1");
					break;
				case "criteria_f1_3":
					driver.switchTo().frame("criteria_f1");
					break;
				case "Past Encounters":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frmTest");
					driver.switchTo().frame(0);
					break;
					
									
				case "externalOrdersResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameExternalOrders");
					driver.switchTo().frame("externalOrdersResult");
					break;
				
					
				case "externalOrdersHeader":
				    driver.switchTo().frame("content");
				    driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameExternalOrders");
					driver.switchTo().frame("externalOrdersHeader");
				    
					
				case "multiPatientOrdersResultingHdr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("multiPatientOrdersResultingHdr");
					break;
				case "multiPatientOrdersResultingHdrResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("multiPatientOrdersResultingHdrResult");
					break;
				case "multiPatientOrdersResultingHdrButtons":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("multiPatientOrdersResultingHdrButtons");
					break;
				case "frameMultiPatientOrdersHdr":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameMultiPatientOrdersHdr");
					break;
				case "frameMultiPatientOrdersHdrDtl":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameMultiPatientOrdersHdrDtl");
					break;
				case "frameMultiPatientOrdersTool":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frameMultiPatientOrdersTool");
					break;
				case "framePatOrderByPrivRelnSearch":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("framePatOrderByPrivRelnSearch");
					break;
				case "framePatOrderByPrResult":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("framePatOrderByPrResult");
					break;
				case "TaskListLeftPatientSearchFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("TaskListLeftFrame");
					driver.switchTo().frame("TaskListLeftPatientSearchFrame");
					break;
				case "TaskListRightFilterFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("TaskListRightFrame");
					driver.switchTo().frame("TaskListRightFilterFrame");
					break;
				
				case "searchFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("searchFrame");
                    break;
                    
				case "volumeFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("volumeFrame");
                    break;
                    
				case "searchframe1":
					driver.switchTo().frame("TermCodeSearchByFrame");
                    driver.switchTo().frame("code_description");
                    driver.switchTo().frame("searchframe");
                    break;
                    
				case "searchResultFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("searchResultFrame");
                    break;

                    
				case "Select_frame1":
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("Select_frame");
                    break;
               
			

				case "Headerpage":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("Headerpage");
					break;
					
				case "SignNotesCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("SignNotesCriteriaFrame");
					break;
					
				case "SignNotesDetailsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("SignNotesDetailsFrame");
					break;
				
				case "RecClinicalNotesSearchToolbarFrame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesSearchToolbarFrame");
				}
				catch(Exception e){
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
				}
					break;
					
				
					
				case "submitframe":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("submitframe");
					break;
					
				case "reaction_view":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("reaction_view");
					break;
			
					
				case "f_search":
					try{
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("f_search");
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("f_tab_frames");
						driver.switchTo().frame("f_search");
					}
					
				break;

				case "f_bed_patient":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_patient_details");
					driver.switchTo().frame("f_bed_patient");
				break;
				case "f_bed_patient1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_patient_details");
					driver.switchTo().frame("f_bed_patient");
				break;
				
				case "f_admin_chart":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_admin_chart");
				break;
				
case "f_admin_chart1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_admin_chart");
				break;
				
				case "RecClinicalNotesSectionFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSectionFrame");
				break;
				case "ViewClinicalNoteCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ViewClinicalNoteCriteriaFrame");
				break;
				case "Events in last 24 hrs":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frmTest");
					driver.switchTo().frame(0);
					break;
				case "Events in last 24 hrs1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryEventsLast24hrs.jsp')]")));
					break;
					
	case "Clinical Notes":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryClinicalNotesD.jsp')]")));
					break;
				case "Past Encounters1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryEncounterList.jsp')]")));
					break;
							
				case "frmTest":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("frmTest");
					break;
				case "frmTest2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryActiveProblemsD.jsp')]")));
					break;
				case "frmTest3":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryClinicalNotesD.jsp')]")));
					break;
				
				case "ChartRecordingListFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ChartRecordingListFrame");
					break;
				case "frmTest1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame(1);
					break;
					
				case "RecClinicalNotesTemplateFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					
					  wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("RecClinicalNotesTemplateFrame"));
					break;
					
				case "RecClinicalNotesSecDetailsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					break;
				case "RecClinicalNotesTemplateFrame1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("RecClinicalNotesTemplateFrame"));
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("C_ALT0000000000036"));
					break;
						
				
				case "RecClinicalNotesTemplateFrame3":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesHeaderFrame");
					driver.switchTo().frame("EditorTitleFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					driver.switchTo().frame("RecClinicalNotesTemplateFrame");
					break;
					
				case "RecClinicalNotesTemplateFrame4":
					//driver.switchTo().frame("EditorTitleFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					driver.switchTo().frame("RecClinicalNotesTemplateFrame");
					break;

				case "RTEditor01":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("f_query_add_mod");
					//WebElement id=driver.findElement(By.name("editor"));
					driver.switchTo().frame("editor");
					break;
				
				case "RTEditor0_1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsEditorFrame");
					driver.switchTo().frame("RecClinicalNotesRTEditorFrame");
					WebElement f= driver.findElement(By.id("RTEditor0"));
					f.sendKeys(Parameter1);
					break;
					
				case "RTEditor0_4":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					WebElement f3= driver.findElement(By.id("RTEditor0"));
					f3.sendKeys(Parameter1);
					break;
					
				case "RTEditor0_3":
					driver.switchTo().frame("editor");
					
					break;
				case "RecClinicalNotesHeaderFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesHeaderFrame");
					break;
				case "patientGenogramTranBtn":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("genogramTranFrame");
					driver.switchTo().frame("patientGenogramTranBtn");
					break;
				case "refusal_main":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("refusal_main");
				break;
				case "refusal_main_result":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("refusal_main_result");
				break;
				case "imageOrderCategory":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("displayTransaction");
					driver.switchTo().frame("imageOrderCategory");
					break;
					
				case "displayImage":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("displayImage");
					break;
					
				case "imageOrderCatalogs":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("orderDetailFrame");
					driver.switchTo().frame("mainFrame");
					driver.switchTo().frame("DetailFrame");
					driver.switchTo().frame("displayTransaction");
					driver.switchTo().frame("imageOrderCatalogs");
					break;
				case "record_details_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("at_tab_frame");
					driver.switchTo().frame("record_details_frame");
					break;
				case "header_frame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("header_frame");
					}
					catch(Exception e)
					{
						driver.switchTo().frame("header_frame");
					}
					break;
				case "complaintsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("meal_tab_frame");
					driver.switchTo().frame("complaintsFrame");
					break;
					
				case "generateMealPlanFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("at_tab_frame");
					driver.switchTo().frame("record_details_frame");
				break;
				case "concl_episode_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("concl_episode_criteria");
				break;
				
				case "group_head":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("group_head");
					}
					catch(Exception e)
					{
						driver.switchTo().frame("group_head");
					}
				break;
				
				case "main_comp":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("main_comp");
				break;
				
				case "group_head1":
					driver.switchTo().frame("TermCodeSearchByFrame");
					driver.switchTo().frame("group_head");
				break;
				
				case "tab_comp":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("tab_comp");
				break;
				
				case "criteria_frame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("criteria_frame");
					}
					catch(Exception e)
					{
						driver.switchTo().frame("criteria_frame");
					}
				break;
				
				case "criteria_frame1":
					driver.switchTo().frame("criteria_frame");
					break;
					
				case "criteria_frame2":
					driver.switchTo().frame("content");
					driver.switchTo().frame("criteria_frame");
					break;
				
				case "result_frame1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("result_frame");
					break;
				case "result_frame3":
					driver.switchTo().frame("content");
					driver.switchTo().frame("query_criteria_frame");
					driver.switchTo().frame("result_frame");
					driver.switchTo().frame("result_frame");
					break;
				case "result_frame2":
					driver.switchTo().frame("result_frame");
					break;
				case "result_frame":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("result_frame");
					}
					catch(Exception e)
					{
						driver.switchTo().frame("result_frame");
					}
				break;
		case "f_query_add_mod_button":
					try{
						driver.switchTo().frame("content");
						driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("f_query_add_mod_button");
						}
					catch(Exception e)
					{
						driver.switchTo().frame("f_query_add_mod_button");
					}
				break;
				case "dispimg_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("dispimg_frame");
				break;
				case "criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("criteria");
				break;
				case "criteria2":
					try{
					driver.switchTo().frame("content");
					driver.switchTo().frame("criteria");	
					break;
					}
					catch(Exception e)
					{
						driver.switchTo().defaultContent();
						driver.switchTo().frame("criteria");
						break;
					}
					
					
				case "Criteria3":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("Criteria");
				break;
				
				
				case "DisplayCriteria2":
					driver.switchTo().frame("details");
					driver.switchTo().frame("details_f2");
					driver.switchTo().frame("DisplayResult");
					driver.switchTo().frame("DisplayCriteria");
					break;
					
				case "ReleaseBed_main":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ReleaseBed_main");
				break;
				case "ReleaseBed_details":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("ReleaseBed_details");
				break;
				case "AEMPSearchCriteriaFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("AEMPSearchCriteriaFrame");
				break;
				case "AEMPSearchResultFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("AEMPSearchResultFrame");
				break;
				case "newbornmain_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newbornmain_frame");
				break;
				
				case "newbornheader_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newbornheader_frame");
				break;
				
				case "newborndtls_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newborndtls_frame");
				break;
			
				case "newbornheader_frame1":
					try{
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newbornheader_frame");
					break;
					}
				catch(Exception e)
					{
					 driver.switchTo().frame("content");
					 driver.switchTo().frame("workAreaFrame");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("newbornheader_frame");
		                 break;  
					}
					
		                 
				case "newborndtls_frame1":
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newborndtls_frame");
				break;
				
				case "newborndtls_frame2":
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newbornmain_frame");
					driver.switchTo().frame("newborndtls_frame");
				break;
				
				case "qa_query_result_tail":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workListResult");
					driver.switchTo().frame("qa_query_result_tail");
				break;
				case "qa_query_result_tail1":
					driver.switchTo().frame("content");
					driver.switchTo().frame("placeMealOrderOPResult");
					driver.switchTo().frame("qa_query_result_tail");
				break;
				case "mealOrderFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("details_frame");
                    break;
                    
                   case "documentMealServedFrame" :
                          driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("details_frame");
                    break;
                    
                   case "meal_plan_details_frame" :
                          driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("meal_tab_frame");
                    driver.switchTo().frame("meal_plan_details_frame");
                    break;
                    
                   case "workListCriteria" :
                          driver.switchTo().frame("content");
                    driver.switchTo().frame("workListCriteria");
                    break;
                      
                    
                   case "qa_query_result":
                	   try{
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("workListResult");
                          driver.switchTo().frame("qa_query_result");
                	   }
                	   catch(Exception e)
                	   {
                		   driver.switchTo().defaultContent();
                		   driver.switchTo().frame("content");
                    	   driver.switchTo().frame("f_query_add_mod");
                           driver.switchTo().frame("qa_query_result");
                           driver.switchTo().frame("qa_query_result");
                	   }
                    break;
                   case "qa_query_result1":
                	 
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("placeMealOrderOPResult");
                          driver.switchTo().frame("qa_query_result");
                     
                  
                   case "AECancelCkoutTrans" :
                       driver.switchTo().frame("content");
                       driver.switchTo().frame("AECancelCkoutTrans");
                       break;
                       
                       
                   case "AECancelCkoutTrans_result" :
                       driver.switchTo().frame("content");
                       driver.switchTo().frame("AECancelCkoutTrans_result");
                       break;
                    
                   case "MealCategoryTab" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("MealCategoryTab");
                    break;
                    
                   case "query_criteria" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("query_criteria");
                    break;
                    
                   case "query_result" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("query_result");
                    break;
                    
                   case "f_query_rep" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("f_query_rep");
                        break;
                 
                   case "at_frame" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("at_frame");
                        break;
                        
                   case "RecordCriteria" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("RecordCriteria");
                        break;
                    
                   case "placeMealOrderOPSearch" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("placeMealOrderOPSearch");
                        break;       
                    
                   case "details_frame" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("details_frame");
                    break;
                  
                   case "tab_frame" :
                          driver.switchTo().frame("content");
                          driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("tab_frame");
                    break;
                    
                   case "tab_frame1" :
                       driver.switchTo().frame("tab_frame");
                       break;
                
                   case "tools" :
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("tools");
                   break;
                   
                   case "f_query_criteria_buttons" :
                	   try{
                		   driver.switchTo().frame("content");
                           driver.switchTo().frame("f_query_add_mod");
                           driver.switchTo().frame("f_query_criteria_buttons");
                	   }
                	   catch(Exception e)
                	   {
                		   driver.switchTo().frame("f_query_add_mod");
                           driver.switchTo().frame("f_query_criteria_buttons");
                	   }
                   break;
                   case "patient_id_display" :
                	   try{
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("patient_id_display");
                	   }
                	   catch(Exception e)
                	   {
                		   driver.switchTo().frame("content");
                           driver.switchTo().frame("f_query_add_mod");
                           driver.switchTo().frame("f_query_result");
                           driver.switchTo().frame("DispMedicationPatFrame");
                           driver.switchTo().frame("patient_id_display");
                	   }
                   break;
                   case "patient_id_display1" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("f_query_result");
                       driver.switchTo().frame("DispMedicationPatFrame");
                       driver.switchTo().frame("patient_id_display");
                       break;
                   case "f_disp_medication_all_stages_legends" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("f_query_result");
                       driver.switchTo().frame("DispMedicationPatDetFrame_3");
                       driver.switchTo().frame("f_disp_medication_all_stages_legends");
                   break;
                   case "f_disp_medication_header":
                       driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("f_query_result");
                       driver.switchTo().frame("DispMedicationPatDetFrame_3");
                       driver.switchTo().frame("f_disp_medication_header");
                  break;
                   case "schdule_dtl":
                	   try{
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("f_query_result");
                       driver.switchTo().frame("schdule_dtl");
                	   }
                	   catch(Exception e)
                	   {
                		   try{
                		   driver.switchTo().defaultContent();
                		   driver.switchTo().frame("content");
                           driver.switchTo().frame("f_query_add_mod");
                		   driver.switchTo().frame("qa_query_result");
                		   driver.switchTo().frame("schdule_dtl");
                		   }
                		   catch(Exception e1)
                		   {
                			   driver.switchTo().defaultContent();
                			   driver.switchTo().frame("content");
                               driver.switchTo().frame("f_query_add_mod");
                               driver.switchTo().frame("f_query_add_mod");
                               driver.switchTo().frame("qa_query_result");
                               driver.switchTo().frame("schdule_dtl");  
                		   }
                	   }
                   break;
                   
                   case "RecDiagnosisAddModify" :
                	   driver.switchTo().frame("RecDiagnosisMain");
                       driver.switchTo().frame("RecDiagnosisAddModify");
                   break;
                 
                   
                   case "RecDiagnosisOpernToolbar" :
                	   driver.switchTo().frame("RecDiagnosisMain");
                       driver.switchTo().frame("RecDiagnosisOpernToolbar");
                   break;
                   case "message_search_frame" :
                	   driver.switchTo().frame("message_search_frame");
                	   driver.switchTo().frame("detailframe");
                       driver.switchTo().frame("message_search_frame");
                   break;
                   case "f_sch_cases_frame" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                	   driver.switchTo().frame("f_tab_frames");
                	   driver.switchTo().frame("f_query_add_mod");
                	   driver.switchTo().frame("f_sch_cases_frame");
                   break;
                   case "detailframe" :
                	   driver.switchTo().frame("message_search_frame");
                       driver.switchTo().frame("detailframe");
                   break;
                   case "frameSalesHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameSalesHeader");
                   break;
                   case "frameSalesListHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameSalesList");
                       driver.switchTo().frame("frameSalesListHeader");
                   break;
                   case "frameSalesListDetail" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameSalesList");
                       driver.switchTo().frame("frameSalesListDetail");
                   break;
                   case "frameSalesReturnHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameSalesReturnHeader");
                   break;
                   case "frameGoodsReceivedNoteHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameGoodsReceivedNoteHeader");
                   break;
                 
                   case "frameGoodsReceivedNoteListHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameGoodsReceivedNoteList");
                       driver.switchTo().frame("frameGoodsReceivedNoteListHeader");
                   break;
                   case "frameGoodsReceivedNoteListDetail" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameGoodsReceivedNoteList");
                       driver.switchTo().frame("frameGoodsReceivedNoteListDetail");
                   break;
                   
                   case "frameGoodsReturnToVendorHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameGoodsReturnToVendorHeader");
                       
                   break;
                   case "frameGoodsReturnToVendorListHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameGoodsReturnToVendorList");
                       driver.switchTo().frame("frameGoodsReturnToVendorListHeader");
                   break;
                   case "frameGoodsReturnToVendorListDetail" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameGoodsReturnToVendorList");
                       driver.switchTo().frame("frameGoodsReturnToVendorListDetail");
                   break;
                   case "frameAdjustStockHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameAdjustStockHeader");
                    break;
                   case "frameAdjustStockListHeader" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameAdjustStockList");
                       driver.switchTo().frame("frameAdjustStockListHeader");
                   break;
                   case "frameAdjustStockListDetail" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameAdjustStockList");
                       driver.switchTo().frame("frameAdjustStockListDetail");
                   break;
                   case "RequestHeaderFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("RequestHeaderFrame");
                   break;
                   case "RequestDetailFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("RequestDetailFrame");
                   break;
                   case "CancelRequestQueryHeaderFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("CancelRequestQueryHeaderFrame");
                   break;
                   
                   case "RequestQueryHeader" :
                	driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                      driver.switchTo().frame("RequestQueryHeader");
                   break;
                   
                     case "CancelRequestQueryResultFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("CancelRequestQueryResultFrame");
                   break;
                   case "reg_prescriptions_query" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("reg_prescriptions_query");
                       break;
        
                   case "criteria0" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("criteria0");
                       break;
                   case "criteria0_1" :
                	    driver.switchTo().frame("criteria0");
                       break;
                
                   case "header" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("header");
                       break;
                       
                   case "Header" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("Header");
                       break;
                       
                   case "header1" :
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("header");
                       break;
                       
                   case "reason_code_top" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("reason_code_top");
                       break;
                       
                   case "reason_code_bottom" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("reason_code_bottom");
                       break;
                       
                   case "patientClassFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("patientClassFrame");
                       break;
                       
                   case "practitionerFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("practitionerFrame");
                       break;
                       
                   case "placeOrderFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("placeOrderFrame");
                       break;
                       
                   case "ordering_rule_top" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("ordering_rule_top");
                       break;
                       	
                   case "ordering_rule_bottom" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("ordering_rule_bottom");
                       break;	
                            
                   case "user_for_review_top1" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("user_for_review_top1");
                       break;
                       
                   case "user_for_review_top2" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("user_for_review_top2");
                       break;
                       
                   case "reason_for_review_bottom" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("reason_for_review_bottom");
                       break;
                       
                   case "user_for_reporting_top" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("user_for_reporting_top");
                       break;
                       
                   case "user_for_reporting_bottom" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("user_for_reporting_bottom");
                       break;
                       
                       
                   case "UserForStoreHeaderFrame" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("UserForStoreHeaderFrame");
                       break;
                       
                
                       
                   case "criteria1" :
                	   driver.switchTo().frame("content");
                       driver.switchTo().frame("criteria");
                       break; 
               
                   case "add_mod":
   					driver.switchTo().frame("content");
   					driver.switchTo().frame("f_query_add_mod");
   					driver.switchTo().frame("add_mod");
   			break;
                  	case "list":
   					driver.switchTo().frame("content");
   					driver.switchTo().frame("f_query_add_mod");
   					driver.switchTo().frame("list"); 
   					break; 
                   case "frameStockItemConsumptionListHeader" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                	   driver.switchTo().frame("frameStockItemConsumptionList");
                       driver.switchTo().frame("frameStockItemConsumptionListHeader");
                       break;  
                   case "TransactionRemarksHeaderFrame" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("TransactionRemarksHeaderFrame");
                       break;     
                   case "IssueQueryHeader" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("IssueQueryHeader");
                       break; 
                   case "IssueQueryResult" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("IssueQueryResult");
                       break; 
case "frameIssueHeader" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                       driver.switchTo().frame("frameIssueHeader"); 
                       break; 
                   case "frameIssueListHeader" :
                	   driver.switchTo().frame("content");
                	   driver.switchTo().frame("f_query_add_mod");
                	   driver.switchTo().frame("frameIssueList");
                       driver.switchTo().frame("frameIssueListHeader");  
                       break;
                       
                       case "eventDetailsFrame" :
                    	   driver.switchTo().frame("content");
                    	   driver.switchTo().frame("workAreaFrame");
                    	   driver.switchTo().frame("ClinicialEventHistoryDetailsFrame");
                           driver.switchTo().frame("eventDetailsFrame");  
                           break;
                           
                       case "eventDataFrame" :
                    	   driver.switchTo().frame("content");
                    	   driver.switchTo().frame("workAreaFrame");
                    	   driver.switchTo().frame("ClinicialEventHistoryDetailsFrame");
                           driver.switchTo().frame("eventDataFrame");  
                           break;
                           
                       case "f_query_result4" :
                    	   driver.switchTo().frame("content");
                    	   driver.switchTo().frame("workAreaFrame");
                    	   driver.switchTo().frame("ClinicialEventHistoryDetailsFrame");
                    	   driver.switchTo().frame("eventDataFrame");
                           driver.switchTo().frame("f_query_result");  
                           break;
                           
                       case "f_tabs_frame" :
                    	   driver.switchTo().frame("content");
                    	   driver.switchTo().frame("workAreaFrame");
                    	   driver.switchTo().frame("ClinicialEventHistoryDetailsFrame");
                    	   driver.switchTo().frame("eventDataFrame");
                           driver.switchTo().frame("f_tabs_frame");  
                           break;
                           
                   case "patEncBannerHdrFrame":
   					driver.switchTo().frame("content");
   					driver.switchTo().frame("patEncBannerHdrFrame");
   					break;
   				  case "patEncBannerDetailsFrame":
   					driver.switchTo().frame("content");
   					driver.switchTo().frame("patEncBannerDetailsFrame");
   					break;
   
                case "detailUpper" :
              	   driver.switchTo().frame("content");
              	   driver.switchTo().frame("f_query_add_mod");
              	   driver.switchTo().frame("detail");
                    driver.switchTo().frame("detailUpper"); 
                     break;
                     
                case "detailUpper1" :
               	   driver.switchTo().frame("f_query_add_mod");
               	   driver.switchTo().frame("detail");
                     driver.switchTo().frame("detailUpper"); 
                      break;
                      
                case "RTEditor0_2" :
               	   driver.switchTo().frame("content");
               	   driver.switchTo().frame("f_query_add_mod");
               	   driver.switchTo().frame("detail");
                   driver.switchTo().frame("detailUpper");
                   Thread.sleep(1000);
                   WebElement ele=driver.findElement(By.tagName("iframe"));
                   System.out.println(ele.getAttribute("name"));
                   ele.sendKeys(Parameter1);
                  
                   break;
                      
                case "detailLower" :
               	   driver.switchTo().frame("content");
               	   driver.switchTo().frame("f_query_add_mod");
               	   driver.switchTo().frame("detail");
                     driver.switchTo().frame("detailLower"); 
                      break;
      
		
				case "frameAcknowledgeHeader" :
             	   driver.switchTo().frame("content");
             	   driver.switchTo().frame("f_query_add_mod");
                   driver.switchTo().frame("frameAcknowledgeHeader"); 
                    break;  
             case "frameAcknowledgeDetail" :
          	   driver.switchTo().frame("content");
          	   driver.switchTo().frame("f_query_add_mod");
                driver.switchTo().frame("frameAcknowledgeDetail"); 
                 break;   
             case "frameAcknowledgeList" :
           	   driver.switchTo().frame("content");
           	   driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("frameAcknowledgeList"); 
                  break;
             case "tab" :
             	   driver.switchTo().frame("content");
             	   driver.switchTo().frame("f_query_add_mod");
                   driver.switchTo().frame("tab"); 
                    break;
                    
             case "tab1" :	
           	   driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("tab"); 
                  break;
                  
             case "Transfer_frame_New":
					driver.switchTo().frame(2);					
					break; 
					
             
             case "frameStockItemConsumptionHeader" :
             	   driver.switchTo().frame("content");
             	   driver.switchTo().frame("f_query_add_mod");
                   driver.switchTo().frame("frameStockItemConsumptionHeader"); 
                    break;
            case "frameBatchSearchQueryResult" :
              	   driver.switchTo().defaultContent();
                    driver.switchTo().frame("frameBatchSearchQueryResult"); 
                     break;  
              case "frameBatchSearchQueryCriteria" :
           	   driver.switchTo().defaultContent();
                 driver.switchTo().frame("frameBatchSearchQueryCriteria"); 
                  break;  
              case "frameStockItemConsumptionListDetail" :
              	driver.switchTo().frame("content");
               	driver.switchTo().frame("f_query_add_mod");
               	driver.switchTo().frame("frameStockItemConsumptionList");
                  driver.switchTo().frame("frameStockItemConsumptionListDetail"); 
                   break; 
                   
              case "MainFrame1":
					driver.switchTo().frame("MainFrame1");
					break;
case "f_search1" :
               	   driver.switchTo().frame("content");
               	   driver.switchTo().frame("f_search");
                     break; 
                     
                case "f_bedheader" :
               	   driver.switchTo().frame("content");
               	   driver.switchTo().frame("workAreaFrame");
               	   driver.switchTo().frame("f_query_add_mod");
               	   driver.switchTo().frame("f_patient_details");
                   driver.switchTo().frame("f_bedheader"); 
                break;
                
		    case "frame0":
				driver.switchTo().frame("frame0");
				break;
				
		      case "Pat_Search_Toolbar_Frame":
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("Pat_Search_Toolbar_Frame"));
				break;      
		    		
         case "Pat_Search_Criteria_Frame":
				driver.switchTo().frame("Pat_Search_Criteria_Frame");
				break;
case "RecClinicalNotesLinkHistNoteCriteriaFrame":
				driver.switchTo().frame("RecClinicalNotesLinkHistNoteCriteriaFrame");
				break;
           case "RecClinicalNotesLinkHistNoteTreeFrame":
				driver.switchTo().frame("RecClinicalNotesLinkHistNoteTreeFrame");
				break;
           case "RecClinicalNotesLinkHistNoteAddButtonsFrame":
				driver.switchTo().frame("RecClinicalNotesLinkHistNoteAddButtonsFrame");
				break;
           case "RecClinicalNotesLinkHistNoteSelectButtons":
				driver.switchTo().frame("RecClinicalNotesLinkHistNoteSelectButtons");
				break;
         case "CAHealthRiskAssessmentOrderCatalogFrm1":
				driver.switchTo().frame("CAHealthRiskAssessmentOrderCatalogFrm");
				break;
			
			case "f_admin_button" :
           	   driver.switchTo().frame("content");
           	   driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("f_admin_button"); 
                 break;
			case "f_admin_button1" :
	           	   driver.switchTo().frame("content");
	           	   driver.switchTo().frame("workAreaFrame");
	           	   driver.switchTo().frame("f_query_add_mod");
	                 driver.switchTo().frame("f_admin_button"); 
	                 break;
			case "f_adrreport_criteria" :
				driver.switchTo().frame("content");
	           	   driver.switchTo().frame("f_query_criteria");
	                 driver.switchTo().frame("f_adrreport_criteria"); 
	                 break; 
			
			
			case "resultFrame" :
				driver.switchTo().frame("content");
	                 driver.switchTo().frame("resultFrame"); 
	                 break;
	                 
			case "resultFrame2" :
				driver.switchTo().frame("content");
				driver.switchTo().frame("f_query_add_mod");
	                 driver.switchTo().frame("resultFrame"); 
	                 break;
	                 
			case "tabsFrame" :
				driver.switchTo().frame("content");
	                 driver.switchTo().frame("tabsFrame"); 
	                 break;
			
			case "f_tpnregimenselect" :
				driver.switchTo().frame("content");
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame("orderDetailFrame");
				driver.switchTo().frame("mainFrame");
				driver.switchTo().frame("DetailFrame");
	           	   driver.switchTo().frame("criteriaPlaceOrderFrame");
	           	 driver.switchTo().frame("f_prescription");
	                 driver.switchTo().frame("f_tpnregimenselect"); 
	                 break;
			case "HeaderFrame" :
				driver.switchTo().frame("content");
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame("orderDetailFrame");
				driver.switchTo().frame("mainFrame");
				driver.switchTo().frame("DetailFrame");
	           	   driver.switchTo().frame("criteriaPlaceOrderFrame");
	           	driver.switchTo().frame("f_prescription");
	           	 driver.switchTo().frame("f_tpnregimen");
	                 driver.switchTo().frame("HeaderFrame"); 
	                 break;
			case "ButtonFrame" :
				driver.switchTo().frame("content");
				driver.switchTo().frame("workAreaFrame");
				driver.switchTo().frame("orderDetailFrame");
				driver.switchTo().frame("mainFrame");
				driver.switchTo().frame("DetailFrame");
	           	   driver.switchTo().frame("criteriaPlaceOrderFrame");
	           	driver.switchTo().frame("f_prescription");
	           	 driver.switchTo().frame("f_tpnregimen");
	                 driver.switchTo().frame("ButtonFrame"); 
	                 break;
	                 
			 case "ButtonFrame1" :
	                 driver.switchTo().frame("ButtonFrame");
	                 break; 
	                 
			 case "AuthorizeAtIssueStoreListFrame" :
          	   driver.switchTo().frame("content");
                 driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("AuthorizeAtIssueStoreListFrame");
                 break;   
			 case "AuthorizeAtIssueStoreDetailFrame" :
	          	   driver.switchTo().frame("content");
	                 driver.switchTo().frame("f_query_add_mod");
	               driver.switchTo().frame("AuthorizeAtIssueStoreDetailFrame");
	                 break;
			 case "f_adrreport_tabdetail" :
	          	   driver.switchTo().frame("content");
	                 driver.switchTo().frame("f_query_criteria");
	               driver.switchTo().frame("f_adrreport_tabdetail");
	                 break;  
			 case "f_tab_detail" :
	          	   driver.switchTo().frame("content");
	                 driver.switchTo().frame("f_query_add_mod");
	               driver.switchTo().frame("f_tab_detail");
	                 break; 
			 case "f_drug_tabs" :
	          	   driver.switchTo().frame("content");
	                 driver.switchTo().frame("f_query_add_mod");
	               driver.switchTo().frame("f_drug_tabs");
	                 break; 
			 
			 case "f_title" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("f_query_add_mod");
	               driver.switchTo().frame("f_tab_detail");
	               driver.switchTo().frame("f_title");
	                 break; 

			 case "patAssessmentAddModifyFrame" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("patAssessmentMainFrame");
	               driver.switchTo().frame("patAssessmentAddModifyFrame");
	                 break;
			 case "patAssessmentDependencyFrame" :
	          	   driver.switchTo().frame("content");
	        	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("patAssessmentMainFrame");
	               driver.switchTo().frame("patAssessmentDependencyFrame");
	                 break;
			
			 case "GeneratePlanProfileDiagnosis" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("GeneratePlanProfile");
	               driver.switchTo().frame("GeneratePlanProfileDiagnosis");
	                 break;
			 case "GeneratePlanTypeBtn" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("GeneratePlanType");
	               driver.switchTo().frame("GeneratePlanTypeBtn");
	                 break;
			 case "GeneratePlanTypeTop" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("GeneratePlanType");
	               driver.switchTo().frame("GeneratePlanTypeTop");
	                 break;
			 case "GeneratePlanDetail" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("GeneratePlanType");
	               driver.switchTo().frame("GeneratePlanDetail");
	                 break;
			 case "patAssessmentButtonsFrame" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	          	   driver.switchTo().frame("patAssessmentMainFrame");
	               driver.switchTo().frame("patAssessmentButtonsFrame");
	                 break;
			 case "matcycleframe" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	               driver.switchTo().frame("matcycleframe");
	                 break;
			 case "maternitytreeframe" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("matFrame");
	               driver.switchTo().frame("maternitytreeframe");
	                 break;
			 case "recMatDeliveryDetails" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("matFrame");
	               driver.switchTo().frame("recMatDeliveryDetails");
	                 break;
			 case "recMatDeliveryDetails1" :
	          	   driver.switchTo().frame("content");
	          	   driver.switchTo().frame("workAreaFrame");
	               driver.switchTo().frame("recMatDeliveryDetails");
	                 break;
			 case "recMatConsBtnsFrame" :
				 driver.switchTo().frame("content");
				 driver.switchTo().frame("matFrame");
	          	   driver.switchTo().frame("maternitytreeframe");
	          	   driver.switchTo().frame("workAreaFrame");
	               driver.switchTo().frame("recMatConsBtnsFrame");
	                 break;

			 case "newbornmain_frame1" :
				 try
				 {
				 driver.switchTo().frame("content");
				 driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("newbornmain_frame");
	                 break; 
				 }
                 catch(Exception e)
                 {
                	 driver.switchTo().defaultContent();
				driver.switchTo().frame("f_query_add_mod");
				driver.switchTo().frame("newbornmain_frame");
			break;
                 }
case "wardacknowledgequeryframe":       
                                  driver.switchTo().frame("content");
                                  driver.switchTo().frame("f_query_criteria");
                                  driver.switchTo().frame("wardacknowledgequeryframe");
                                  break;
                                  
case "queryFrame":       
    			driver.switchTo().frame("content");
    			driver.switchTo().frame("f_query_add_mod");
    			driver.switchTo().frame("queryFrame");
    break;
    
case "buttons":       
	driver.switchTo().frame("content");
	driver.switchTo().frame("f_query_add_mod");
	driver.switchTo().frame("buttons");
break;

case "facilityForUserTopFrame":
	driver.switchTo().frame("content");
	driver.switchTo().frame("f_query_add_mod");
	driver.switchTo().frame("facilityForUserTopFrame");
break;
    
case "ExistingOrderSearch4":
                                  driver.switchTo().frame("workAreaFrame");
                                  driver.switchTo().frame("orderDetailFrame");
                                  driver.switchTo().frame("ExistingOrderSearch");
                                  break;

case "AssignRelationshipFrame1":
              driver.switchTo().frame("AssignRelationshipFrame");
              break;
case "ResultEntryDtl":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ResultEntryDtl");
					break; 
case "ResultEntryReport":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("ResultEntryReport");
					break;
			 case "eval_cp_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("eval_cp_criteria");
					break;
			 case "eval_cp_result":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("eval_cp_result");
					break;
 case "SignNotesToolsFrame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("SignNotesToolsFrame");
					break;
case "SpecimenOrderBtn":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("SpecimenOrderBtn");
					break;	
case "transfer_from":
					driver.switchTo().frame("content");
					driver.switchTo().frame("transfer_from");
					break;
case "transfer_to":
					driver.switchTo().frame("content");
					driver.switchTo().frame("transfer_to");
					break;
 case "transfer_criteria":
					driver.switchTo().frame("content");
					driver.switchTo().frame("transfer_criteria");
					break;
	case "frameStockTransferListHeader":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameStockTransferList");
					driver.switchTo().frame("frameStockTransferListHeader");
					break;
			 case "frameStockTransferListDetail":
					driver.switchTo().frame("content");
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("frameStockTransferList");
					driver.switchTo().frame("frameStockTransferListDetail");
					break;	
                case "searchFrame1":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("searchFrame");
                    break;
 	case "search_collapse_frame":
					driver.switchTo().frame("content");
					driver.switchTo().frame("search_collapse_frame");
					break;	
			 case "query_search_result":
					driver.switchTo().frame("content");
					driver.switchTo().frame("query_search_result");
					break;	
case "CASectionTemplateToolbarFrame1": 
                                  driver.switchTo().frame("content");
                                  driver.switchTo().frame("f_query_add_mod");
                                  driver.switchTo().frame("CASectionTemplateToolbarFrame");
                                  break;

   case "RecPatChiefComplaintAddModifyFrame1":
						driver.switchTo().frame("RecPatChiefComplaintAddModifyFrame");
						break;
case "submitframe2":
driver.switchTo().frame("submitframe");
						break;
case "RecPatChiefComplaintDiagLookupResultFrame":
						driver.switchTo().frame("RecPatChiefComplaintDiagLookupCriteriaFrame");
						driver.switchTo().frame("RecPatChiefComplaintDiagLookupResultFrame");
						break;
		
case "searchCriteria1":
	driver.switchTo().frame("searchCriteria");
	break;

 case "MedRepAuditCriteria":
						driver.switchTo().frame("content");
						driver.switchTo().frame("MedRepAuditCriteria");
						break;	
				 case "MedRepAuditResult":
						driver.switchTo().frame("content");
						driver.switchTo().frame("MedRepAuditResult");
						break;
   case "Pline_frame": // Added in 10.2.3 
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("Pline_frame");
                    break;
                     case "patientLine": // Added in 10.2.3 
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("patientLine");
                    break;
                     case "patLine": // Added in 10.2.3      
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("patLine");
                    break;

	case "f_query_add_mod7":
                                driver.switchTo().frame("content");
                                driver.switchTo().frame(2);
                                break; 
case "practitioner_main":
                    driver.switchTo().frame("content");    
                    driver.switchTo().frame("f_query_add_mod");
                    driver.switchTo().frame("practitioner_main");
                    break;

case "CAPrintingEmrPatientResult":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("CAPrintingEmrPatientResult");
             break;
       case "CAPrintingEmrPatientCriteria":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("CAPrintingEmrPatientCriteria");
             break;
              case "CAPrintingEmrPatientSubRes":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("CAPrintingEmrPatientSubRes");
             break;

case "OPCancelChkout_Result":
       driver.switchTo().frame("content");
       driver.switchTo().frame("OPCancelChkout_Result");
       break;

case "OPCancelChkout_Search":
driver.switchTo().frame("content");
driver.switchTo().frame("OPCancelChkout_Search");
break;

case "DisplayCriteria4":
                                  driver.switchTo().frame("content");
                                  driver.switchTo().frame("workAreaFrame");
                                  driver.switchTo().frame("details");
                                  driver.switchTo().frame("details_f2");
                                  driver.switchTo().frame("DisplayResult");
                                  driver.switchTo().frame(1);
case "DisplayCriteria5":
                                  driver.switchTo().frame("content");
                                  driver.switchTo().frame("workAreaFrame");
                                  driver.switchTo().frame("details");
                                  driver.switchTo().frame("details_f2");
                                  driver.switchTo().frame("DisplayResult");
                                  driver.switchTo().frame("DisplayCriteria");
                                  break;


					case "detailframe1":  
                         driver.switchTo().frame("message_search_frame");
                         driver.switchTo().frame(0);
                         driver.switchTo().frame("detailframe");
                         break;
	
                     case "DetailFrame1":  
                         driver.switchTo().frame("f_query_add_mod");
                         driver.switchTo().frame("ChkListDetailFrame");
                         driver.switchTo().frame("DetailFrame");
                         break;
                  
                         
                     case "pline1":
                         driver.switchTo().frame("RecDiagnosisMain");
                         driver.switchTo().frame("pline");
                         break;  
                         
          
                     case "RecDiagnosisCurrentDiagheader":
                         driver.switchTo().frame("RecDiagnosisMain");
                         driver.switchTo().frame("RecDiagnosisCurrentDiagheader");
                         break; 
                     case "RecDiagnosisCurrentDiagLegend":
                         driver.switchTo().frame("RecDiagnosisMain");
                         driver.switchTo().frame("RecDiagnosisCurrentDiagLegend");
                         break; 
                     case "RecDiagnosisCurrentDiagLegend1":
                        driver.switchTo().frame("RecDiagnosisCurrentDiagLegend");
                         break; 
                     case "CommonPatientLineFrame":
                         driver.switchTo().frame("patient_title");
                         driver.switchTo().frame("CommonPatientLineFrame");
                         break;   
                     case "f_header1":
                    	 driver.switchTo().frame("patient_title");
                    	 driver.switchTo().frame("CommonPatientLineFrame");
                         driver.switchTo().frame("OTPatientLineFrame");
                         driver.switchTo().frame("f_header");
                         break; 
                     case "OTPatientLineFrame":
                         driver.switchTo().frame("patient_title");
                         driver.switchTo().frame("OTPatientLineFrame");
                         break; 
                     
                         
                     case "AssignRelationshipViewFrame":
                         driver.switchTo().frame("plineFrame");
                         driver.switchTo().frame("AssignRelationshipViewFrame");
                         break;
                     
                     case "messageFrame3":
                         driver.switchTo().frame("AssignRelationshipFrame");
                         driver.switchTo().frame("messageFrame");
                         break;   
                         
                     case "messageFrame4":
                         driver.switchTo().frame("RecDiagnosisMain");
                         driver.switchTo().frame("messageFrame");
                         break;   
                         
                     case "messageFrame5":
                    	 driver.switchTo().frame("content");
                         driver.switchTo().frame("workAreaFrame");
                         driver.switchTo().frame("RecDiagnosisMain");
                         driver.switchTo().frame("messageFrame");
                         break;
                         
                         
                     case "euip_tab":
                         driver.switchTo().frame("RecordFrame");
                         driver.switchTo().frame("euip_tab");
                         break;
                         
                     case "frameSalesHeader1" :
                  	   driver.switchTo().frame("commontoolbarFrame");
                         driver.switchTo().frame("f_query_add_mod");
                         driver.switchTo().frame("frameSalesHeader");
                     break;
                     case "frameSalesListHeader1" :
                  	   driver.switchTo().frame("commontoolbarFrame");
                         driver.switchTo().frame("f_query_add_mod");
                         driver.switchTo().frame("frameSalesList");
                         driver.switchTo().frame("frameSalesListHeader");
                     break;
                     case "frameSalesListDetail1" :
                  	   driver.switchTo().frame("commontoolbarFrame");
                         driver.switchTo().frame("f_query_add_mod");
                         driver.switchTo().frame("frameSalesList");
                         driver.switchTo().frame("frameSalesListDetail");
                     break;
                     case "pre_morbid_record_frame" :
                    	   driver.switchTo().frame("pre_anesth_tab_details_frame");
                           driver.switchTo().frame("pre_morbid_details_tab_frame");
                           driver.switchTo().frame("pre_morbid_record_frame");
                       break;
                    
                     case "EditorTitleFrame" :
                    		driver.switchTo().frame("content");
                  	   driver.switchTo().frame("workAreaFrame");
                         driver.switchTo().frame("RecClinicalNotesFrame");
                         driver.switchTo().frame("EditorTitleFrame");
                     break;
                     
                     case "SearchResultsFrame":
                         driver.switchTo().frame("ResultFrame");
                         driver.switchTo().frame("SearchResultsFrame");
                         break;
                     case "CommonPatientLineFrame1":
                         driver.switchTo().frame("OtPatientLineFrame");
                         driver.switchTo().frame("CommonPatientLineFrame");
                         break;
                         
                     case "untowards_evt_record_frame":
                      driver.switchTo().frame("result_frame");
                      driver.switchTo().frame("untowards_evt_record_frame");
                  break;
                  
                     case "RecPatChiefComplaintDiagLookupResultFrame1":
 						driver.switchTo().frame("RecPatChiefComplaintDiagLookupResultFrame");
 						break;

case "procedurelistquery":                                                                                         
          driver.switchTo().frame("FlowSheet");
          driver.switchTo().frame("procedurelistquery");
          break;

case "procedurelisttitle":                                                                                             
          driver.switchTo().frame("FlowSheet");
          driver.switchTo().frame("procedurelisttitle");
          break;
case "procedurelistresult":                                                                                         
          driver.switchTo().frame("FlowSheet");
          driver.switchTo().frame("procedurelistresult");
          break;
case "procedurelistresultdetail":                                                                                              
          driver.switchTo().frame("FlowSheet");
          driver.switchTo().frame("procedurelistresultdetail");
          break;
case "AddModifypatFinDetailsInsButtonFrame":
          driver.switchTo().frame("MainFrame2");    
          driver.switchTo().frame("InsuranceFrame");
          driver.switchTo().frame("AddModifypatFinDetailsInsButtonFrame");
          break;
case "search_criteria_frame":                                                                                                   
          driver.switchTo().frame("content");
          driver.switchTo().frame("search_criteria_frame");
          break;
case "default_val_frame":                                                                                          
          driver.switchTo().frame("content");
          driver.switchTo().frame("default_val_frame");
          break;
case "IssuesView":                                                                                         
          driver.switchTo().frame("content");
          driver.switchTo().frame("issue_tab");
          driver.switchTo().frame("IssuesView");
          break; 
case "OutstanListDetail":                                                                                             
                  driver.switchTo().frame("content");
                  driver.switchTo().frame("issue_detail");
                  driver.switchTo().frame("OutstanListDetail");
                 break;
case "select_files":                                                                                         
                 driver.switchTo().frame("content");
                 driver.switchTo().frame("select_files");
                 break;                                                                                 
case "Headerpage1":
                  driver.switchTo().frame("Headerpage");
                  break;
case "reaction_view1":
                  driver.switchTo().frame("reaction_view");
                  break;

case "AuthorizeAtIssueStoreHeaderFrame":
          driver.switchTo().frame("content");
          driver.switchTo().frame("f_query_add_mod");
          driver.switchTo().frame("AuthorizeAtIssueStoreHeaderFrame");
          break;
case "frameRTVHistorySearchQueryCriteria":
          driver.switchTo().frame("content");
          driver.switchTo().frame("f_query_add_mod");
          driver.switchTo().frame("frameRTVHistorySearchQueryCriteria");
          break;
case "frameRTVHistoryQueryResult":
          driver.switchTo().frame("content");
          driver.switchTo().frame("f_query_add_mod");
          driver.switchTo().frame("frameRTVHistoryQueryResult");
          break;
case "frameGRNHistorySearchQueryCriteria":
          driver.switchTo().frame("content");
          driver.switchTo().frame("f_query_add_mod");
          driver.switchTo().frame("frameGRNHistorySearchQueryCriteria");
          break;
case "frameGRNHistoryQueryResult":
        driver.switchTo().frame("content");
        driver.switchTo().frame("f_query_add_mod");
        driver.switchTo().frame("frameGRNHistoryQueryResult");
        break;
case "PatientSecondaryTriageDetails":
                    driver.switchTo().frame("PatientSecondaryTriageDetails");
                    driver.switchTo().frame("f_query_add_mod");
                   break;


case "ResearchPatientRecordframe":
driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("ResearchPatientRecordframe");
                    break;

case "ExistScheduleTreeFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("ExistScheduleTreeFrame");
                    break;
    case "VaccineResult":
                                driver.switchTo().frame("content");
                                driver.switchTo().frame("workAreaFrame");
                                driver.switchTo().frame("ExistScheduleDetailFrame");
                                driver.switchTo().frame("VaccineTabResult");
                                driver.switchTo().frame("VaccineResult");
                                break;   

case "VaccinAdminDetailsFrame":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("VaccinationFrame");
                    driver.switchTo().frame("VaccinAdminDetailsFrame");
                    break;

case "image34":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("workAreaFrame");
                    driver.switchTo().frame("dummy");
                    driver.switchTo().frame("image34");
                    break;
case "Result":
                driver.switchTo().frame("content");
                driver.switchTo().frame("workAreaFrame");
                driver.switchTo().frame(1);
            driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryResultsD.jsp')]")));
                break;
case "Result1":
	driver.switchTo().frame("content");
    driver.switchTo().frame("f_query_add_mod");
    driver.switchTo().frame("Result");
    break;

case "PatientCare":
                driver.switchTo().frame("content");
                driver.switchTo().frame("workAreaFrame");
                driver.switchTo().frame(1);
driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryCareFlowsheetD.jsp')]")));
                break;

case "HighRisk":
                driver.switchTo().frame("content");
                driver.switchTo().frame("workAreaFrame");
                driver.switchTo().frame(1);
  driver.switchTo().frame(driver.findElement(By.xpath("//iframe[contains(@src,'ChartSummaryMedicalAlerts.jsp')]")));
                break;
case "button_f1":
                    driver.switchTo().frame("criteria_f0");
                    driver.switchTo().frame("details");
                    driver.switchTo().frame("button_f1");
                    break;

case "DisplayCriteria3":
                driver.switchTo().frame("criteria_f0");
                driver.switchTo().frame("details");
                driver.switchTo().frame("details_f2");
              driver.switchTo().frame("DisplayResult");
                driver.switchTo().frame(0);
              break;
case "view":
                driver.switchTo().frame("PatientSecondaryTriageDetails");
                driver.switchTo().frame("view");
                break;
  case "MainFrame":
               driver.switchTo().frame("PatientSecondaryTriageDetails");
              driver.switchTo().frame("MainFrame");
                break;
case "splChtSpecialViewFrame":
                    driver.switchTo().frame("viewsFrame");
                    driver.switchTo().frame("splChtSpecialViewFrame");
                    break;
case "splChtCustomViewFrame":
                    driver.switchTo().frame("viewsFrame");
                    driver.switchTo().frame("splChtCustomViewFrame");
                    break;
case "RecClinicalNotesSecDetailsFrame1":
     					driver.switchTo().frame("content");
     					driver.switchTo().frame("workAreaFrame");
     					driver.switchTo().frame("RecClinicalNotesFrame");
     					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
     					WebElement f1= driver.findElement(By.id("RTEditor0"));
    					f1.sendKeys(Parameter1);
     					break;
case "RecClinicalNotesSecDetailsFrame2":
		driver.switchTo().frame("content");
		driver.switchTo().frame("workAreaFrame");
		driver.switchTo().frame("RecClinicalNotesFrame");
		driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
		break;


case "f_patientline":
					try{
                    driver.switchTo().frame("f_query_result");
                    driver.switchTo().frame("f_patientline");
             break;
					}
					catch(Exception e){
						driver.switchTo().frame("content");
						driver.switchTo().frame("f_query_add_mod");
						driver.switchTo().frame("f_query_result");
	                    driver.switchTo().frame("f_patientline");
	             break;
					}
case "f_orderdetailsframe":
                    driver.switchTo().frame("f_orderframe");
                    driver.switchTo().frame("f_orderdetailsframe");
                    break;
case "f_ordercolorframe":
                 driver.switchTo().frame("f_orderframe");
                 driver.switchTo().frame("f_ordercolorframe");
          break;
case "details_f1":
                    driver.switchTo().frame("criteria_f0"); 
                    driver.switchTo().frame("details");
                    driver.switchTo().frame("details_f1");
                    break;


case "f_disp_medication_delivery":
				 driver.switchTo().frame("content");
				 driver.switchTo().frame("f_query_add_mod");
				 driver.switchTo().frame("f_query_result");
                 driver.switchTo().frame("DispMedicationPatDetFrame_3");
                 driver.switchTo().frame("f_disp_medication_delivery");
          break;
case "f_query_criteria1":
				 driver.switchTo().frame("content");
                 driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("f_query_criteria");
          break;
case "f_criteria":
				 driver.switchTo().frame("content");
                 driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("f_criteria");
          break;
case "f_result":
				 driver.switchTo().frame("content");
                 driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("f_result");
          break;

case "f_query_result3":
				 driver.switchTo().frame("content");
                 driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("f_query_result");
                 driver.switchTo().frame("f_query_result");
          break;
case "f_enquiry":
				 driver.switchTo().frame("content");
              driver.switchTo().frame("f_query_add_mod");
              driver.switchTo().frame("f_enquiry");
       break;
case "f_patient":
				 driver.switchTo().frame("content");
                 driver.switchTo().frame("f_query_add_mod");
                 driver.switchTo().frame("f_patient");
          break;
case "wardacknowledgeactionframe":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_query_criteria");
                    driver.switchTo().frame("wardacknowledgeactionframe");
                    break;		


				default:
					framehandle1(framename,driver,Parameter1);
							

				}
		}
		catch(Exception e){}
				try{	
				framename=(String) jsExecutor.executeScript("return self.name");
				System.out.println("framename....................."+framename+"\n");
				//return framename;
				}
				catch(Exception e){}
				return framename;
	}
		
	
	public static String framehandle1(String framename,WebDriver driver,String Parameter1) throws InterruptedException
	{
		WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(10));
		
		JavascriptExecutor jsExecutor = (JavascriptExecutor)driver;
		try{
			
				driver.switchTo().defaultContent();
				switch(framename)
				{
				     case "RequestQueryResult" :
	             	 driver.switchTo().frame("content");
	                    driver.switchTo().frame("f_query_add_mod");
	                   driver.switchTo().frame("RequestQueryResult");
	                break;
	                   
	                   case "RequestListFrame" :
	              	  driver.switchTo().frame("content");
	                     driver.switchTo().frame("f_query_add_mod");
	                    driver.switchTo().frame("RequestListFrame");
	                 break;
	                   
	                 case "frameSalesAndReturnHistoryQueryCriteria" :
	               	  driver.switchTo().frame("content");
	                      driver.switchTo().frame("f_query_add_mod");
	                     driver.switchTo().frame("frameSalesAndReturnHistoryQueryCriteria");
	                  break;
	                   
	                 case "frameSalesAndReturnHistoryQueryResult" :
	                	  driver.switchTo().frame("content");
	                       driver.switchTo().frame("f_query_add_mod");
	                      driver.switchTo().frame("frameSalesAndReturnHistoryQueryResult");
	                   break;
	                   
	                 case "query3" :
	                       driver.switchTo().frame("mainFrame");
	                      driver.switchTo().frame("query3");
	                   break;
	                   
	                 case "query2" :
	                       driver.switchTo().frame("mainFrame");
	                      driver.switchTo().frame("query2");
	                   break;
	                   
	                 case "btn_frame" :
	                       driver.switchTo().frame("main_frame");
	                      driver.switchTo().frame("btn_frame");
	                   break;
	                   
	                 case "ItemForStoreTabFrame" :
	                       driver.switchTo().frame("content");
	                       driver.switchTo().frame("f_query_add_mod");
	                      driver.switchTo().frame("ItemForStoreTabFrame");
	                   break;
	                   
	                 case "UserForStoreListFrame" :
	                	   driver.switchTo().frame("content");
	                       driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("UserForStoreListFrame");
	                       break;
	                       
	                 case "TransactionRemarksDetailFrame" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("TransactionRemarksDetailFrame");
	                       break;
	                       
	                 case "CancelAuthQueryHeaderFrame" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("CancelAuthQueryHeaderFrame");
	                       break;
	                       
	                 case "CancelAuthQueryResultFrame" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("CancelAuthQueryResultFrame");
	                       break; 
	                       
	                 case "frameGoodsReceivedNoteDetail" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("frameGoodsReceivedNoteDetail");
	                       break;
	                       
	                 case "GoodsReceivedNoteQueryHeader" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("GoodsReceivedNoteQueryHeader");
	                       break;
	                       
	                 case "GoodsReceivedNoteQueryResult" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("GoodsReceivedNoteQueryResult");
	                       break;
	                       
	                 case "lookup_cancel" :
	                       driver.switchTo().frame("lookup_cancel");
	                       break;
	                       
	                 case "DetailsFrame" :
	                	   driver.switchTo().frame("content");
	                	   driver.switchTo().frame("f_query_add_mod");
	                       driver.switchTo().frame("DetailsFrame");
	                       break;
	                       
	                 case "AddModifypatFinDetailsInsBodyFrame" :
	                	   driver.switchTo().frame("MainFrame2");
	                	   driver.switchTo().frame("InsuranceFrame");
	                       driver.switchTo().frame("AddModifypatFinDetailsInsBodyFrame");
	                       break;    
	                       
	                 case "AddModifypatFinDetailsInsButtonFrame" :
	                	   driver.switchTo().frame("MainFrame2");
	                	   driver.switchTo().frame("InsuranceFrame");
	                       driver.switchTo().frame("AddModifypatFinDetailsInsButtonFrame");
	                       break;
	                 case "f_detail4":
	 					driver.switchTo().frame("orderDetailFrame");
	 					driver.switchTo().frame("mainFrame");
	 					driver.switchTo().frame("DetailFrame");
	 					driver.switchTo().frame("criteriaPlaceOrderFrame");
	 					driver.switchTo().frame("f_detail");
	 					break;
	 					
	                 case "f_bedheader1" :
	                 	   driver.switchTo().frame("content");
	                 	   driver.switchTo().frame("f_query_add_mod");
	                 	   driver.switchTo().frame("f_patient_details");
	                     driver.switchTo().frame("f_bedheader"); 
	                  break;
	                  
	                 case "f_query_add_user" :
	                 	   driver.switchTo().frame("content");
	                     driver.switchTo().frame("f_query_add_user"); 
	                  break;
	                  
	                 case "f_query_add_mod9":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("f_tab_detail");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    break;	

	                	case "f_button4":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("f_tab_detail");
	                	    driver.switchTo().frame("f_button");
	                	    break;	
	                	case "f_add_mod_detail":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("f_tab_detail");
	                	    driver.switchTo().frame("f_add_mod_detail");
	                	    break;	
	                	    
	                	case "PrepareGroupHeaderFrame":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("PrepareGroupHeaderFrame");
	                	    break;
	                	    
	                	case "frameROFEntryHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameROFEntryHeader");
	                	    break;
	                	    
	                	case "frameROFEntryQueryCriteria":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameROFEntryQueryCriteria");
	                	    break;
	                	    
	                	case "frameROFEntryQueryResult":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameROFEntryQueryResult");
	                	    break;
	                	    
	                	case "frameWashingAddModify":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameWashingAddModify");
	                	    break;
	                	    
	                	case "frameWashingDetail":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameWashingDetail");
	                	    break;
	                	    
	                	case "framePackingQueryCriteria":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("framePackingQueryCriteria");
	                	    break;
	                	    
	                	case "framePackingQueryResult":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("framePackingQueryResult");
	                	    break;
	                	    
	                	case "frameAutoclavingHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameAutoclavingHeader");
	                	    break;
	                	    
	                	case "frameAutoclavingList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameAutoclavingList");
	                	    break;
	                	    
	                	case "frameDispatchAddModify":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameDispatchAddModify");
	                	    break;
	                	    
	                	case "frameDispatchDetail":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameDispatchDetail");
	                	    break;
	                	    
	                	case "qryCriteriaTrayDtl":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("qryCriteriaTrayDtl");
	                	    break;
	                	    
	                	case "frameRequestGroupHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameRequestGroupHeader");
	                	    break;
	                	    
	                	case "frameRequestGroupList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameRequestGroupList");
	                	    break;
	                	    
	                	case "frameCancelRequestQueryCriteria":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameCancelRequestQueryCriteria");
	                	    break;
	                	    
	                	case "frameCancelRequestQueryResult":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameCancelRequestQueryResult");
	                	    break;
	                	    
	                	case "frameCancelRequestHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameCancelRequestHeader");
	                	    break;
	                	    
	                	case "frameIssueGroupList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameIssueGroupList");
	                	    break;
	                	    
	                	case "AcknowledgeHeaderframe":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("AcknowledgeHeaderframe");
	                	    break;
	                	    
	                	case "AcknowledgeDetailframe":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("AcknowledgeDetailframe");
	                	    break;
	                	    
	                	case "frameReturnGroupHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameReturnGroupHeader");
	                	    break;
	                	    
	                	case "frameReturnGroupList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameReturnGroupList");
	                	    break;
	                	    
	                	case "frameReturnGroupDetail":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameReturnGroupDetail");
	                	    break;
	                	    
	                	case "frameLoanRequestGroupHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameLoanRequestGroupHeader");
	                	    break;
	                	    
	                	case "frameLoanIssueGroupList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameLoanIssueGroupList");
	                	    break;
	                	    
	                	case "frameLoanReturnGroupHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameLoanReturnGroupHeader");
	                	    break;
	                	    
	                	case "frameLoanReturnGroupList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameLoanReturnGroupList");
	                	    break;
	                	    
	                	case "frameVendorLoanRequestHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameVendorLoanRequestHeader");
	                	    break;
	                	    
	                	case "frameVendorLoanRequestList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameVendorLoanRequestList");
	                	    break;
	                	    
	                	case "frameVendorLoanReturnHeader":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameVendorLoanReturnHeader");
	                	    break;
	                	    
	                	case "frameVendorLoanReturnList":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("frameVendorLoanReturnList");
	                	    break;
	                	    
	                	case "RejectRequestQueryHeaderFrame":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("RejectRequestQueryHeaderFrame");
	                	    break;
	                	    
	                	case "RejectRequestQueryResultFrame":
	                		driver.switchTo().frame("content");
	                	    driver.switchTo().frame("f_query_add_mod");
	                	    driver.switchTo().frame("RejectRequestQueryResultFrame");
	                	    break;
	                	case "Diagnosis":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("workAreaFrame");
	    					driver.switchTo().frame("frmTest");
	    					driver.switchTo().frame(0);
	    					break;
	                	case "ConfigDispCritHdr":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("ConfigDispCritHdr");
	    					break;
	                	case "ConfigDispCritList":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("ConfigDispCritList");
	    					break;
	    					
	                	case "ProblemLSD":
	    					driver.switchTo().frame(driver.findElement(By.xpath("//frame[contains(@src,'../../eCA/jsp/ShowSupportingToolbar.jsp')]")));
	    					break;
	                	case "DiagnosisSD":
	    					driver.switchTo().frame(driver.findElement(By.xpath("//frame[contains(@src,'../../eMR/jsp/ShowExternalCauseToolbar.jsp')]")));
	    					break;
	                	case "imParamFrm":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("imParamFrm");
	    					break;
	                	case "ExistScheduleTreeFrame":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("workAreaFrame");
	    					driver.switchTo().frame("ExistScheduleTreeFrame");
	    					break;
	                	case "VaccinAdminDetailsFrame":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("workAreaFrame");
	    					driver.switchTo().frame("VaccinationFrame");
	    					driver.switchTo().frame("VaccinAdminDetailsFrame");
	    					break;
	                	case "VaccineResult":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("workAreaFrame");
	    					driver.switchTo().frame("ExistScheduleDetailFrame");
	    					driver.switchTo().frame("VaccineTabResult");
	    					driver.switchTo().frame("VaccineResult");
	    					break;
	                	case "PatProvrelationCreteriaFrame":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("workAreaFrame");
	    					driver.switchTo().frame("PatProvrelationCreteriaFrame");
	    					break;
	                	case "PatProvrelationResultFrame":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("workAreaFrame");
	    					driver.switchTo().frame("PatProvrelationResultFrame");
	    					break;
	                	case "addModframe":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("f_query_add_mod");
	    					driver.switchTo().frame("addModframe");
	    					break;
	                	case "selectcriteriaframe":
	    					driver.switchTo().frame("content");
	    					driver.switchTo().frame("f_query_add_mod");
	    					driver.switchTo().frame("selectcriteriaframe");
	    					break;				
	                case "frameHeader":
		    					driver.switchTo().frame("content");
		    					driver.switchTo().frame("f_query_add_mod");
		    					driver.switchTo().frame("frameHeader");
		    					break;
	                case "frameBatchSearchQueryResult1" :
	                     driver.switchTo().frame("frameBatchSearchQueryResult"); 
	                      break; 
	                case "frameStockStatusSearchQueryCriteria":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("frameStockStatusSearchQueryCriteria");
    					break;
	                case "frameStockStatusQueryResult":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("frameStockStatusQueryResult");
    					break;	
	                case "detailframe":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("detailframe");
    					break;
	                case "f_prescription_form1":
	                    driver.switchTo().frame("orderDetailFrame");
	                    driver.switchTo().frame("mainFrame");
	                    driver.switchTo().frame("DetailFrame");
	                    driver.switchTo().frame("criteriaPlaceOrderFrame");
	                    driver.switchTo().frame("f_prescription");
	                    driver.switchTo().frame("f_prescription_form");
	                    break;
	                case "detailframe3":
	                    driver.switchTo().frame("detailframe");
	                    break;
	                case "frameDrugQuotaLimitHeader":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("frameDrugQuotaLimitHeader");
    					break;
	                case "AssignPrivilegeGroupHdr":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("AssignPrivilegeGroupHdr");
    					break;
	                case "AssignPrivilegeGroupDtl":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("AssignPrivilegeGroupDtl");
    					break;
	                case "AssignPrivilegeGroupResult":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("f_query_add_mod");
    					driver.switchTo().frame("AssignPrivilegeGroupResult");
    					break;
	                case "f_query_criteria6":
    					driver.switchTo().frame("f_query_criteria");
    					break;
	                case "QuotaLimitApprovalCriteriaFrame":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("QuotaLimitApprovalCriteriaFrame");
    					break;
	                case "QuotaLimitApprovalBottomRight1":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("QuotaLimitApprovalBottom");
    					driver.switchTo().frame("QuotaLimitApprovalBottomRight");
    					driver.switchTo().frame("QuotaLimitApprovalBottomRight1");
    					break;
	                case "QuotaLimitApprovalBottomRight2":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("QuotaLimitApprovalBottom");
    					driver.switchTo().frame("QuotaLimitApprovalBottomRight");
    					driver.switchTo().frame("QuotaLimitApprovalBottomRight2");
    					break;
	                case "SpecialApprovalOrdersTop":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("workAreaFrame");
    					driver.switchTo().frame("SpecialApprovalOrdersTop");
    					break;
	                case "SpecialApprovalOrdersBottomRight1":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("workAreaFrame");
    					driver.switchTo().frame("SpecialApprovalOrdersBottom");
    					driver.switchTo().frame("SpecialApprovalOrdersBottomRight");
    					driver.switchTo().frame("SpecialApprovalOrdersBottomRight1");
    					break;
	                case "SpecialApprovalOrdersBottomRight2":
    					driver.switchTo().frame("content");
    					driver.switchTo().frame("workAreaFrame");
    					driver.switchTo().frame("SpecialApprovalOrdersBottom");
    					driver.switchTo().frame("SpecialApprovalOrdersBottomRight");
    					driver.switchTo().frame("SpecialApprovalOrdersBottomRight2");
    					break;
	                case "f_price_detail":
	                	driver.switchTo().frame("content");
       					driver.switchTo().frame("f_price_detail");
    					break;
    					
	                case "GROUP_HDR":
	                	driver.switchTo().frame("content");
	                	driver.switchTo().frame("f_query_add_mod");
       					driver.switchTo().frame("GROUP_HDR");
    					break;
    					
	                case "GROUP_DTLS":
	                	driver.switchTo().frame("content");
	                	driver.switchTo().frame("f_query_add_mod");
       					driver.switchTo().frame("GROUP_DTLS");
    					break;
    					
	                case "GROUP_DTLS_TITLE":
	                	driver.switchTo().frame("content");
	                	driver.switchTo().frame("f_query_add_mod");
       					driver.switchTo().frame("GROUP_DTLS_TITLE");
    					break;
    					
	                case "BHT_query":
	                	driver.switchTo().frame("content");
       					driver.switchTo().frame("BHT_query");
    					break;
    					
	                case "BHT_result":
	                	driver.switchTo().frame("content");
       					driver.switchTo().frame("BHT_result");
    					break;
				
					
				case "f_adding_criteria":
                	driver.switchTo().frame("content");
                	driver.switchTo().frame("createFrame");
   					driver.switchTo().frame("f_adding_criteria");
					break;
					
				case "framePhysicalInventoryEntryHeader":
                	driver.switchTo().frame("content");
                	driver.switchTo().frame("f_query_add_mod");
   					driver.switchTo().frame("framePhysicalInventoryEntryHeader");
					break;
					
				case "framePhysicalInventoryDetail":
                	driver.switchTo().frame("content");
                	driver.switchTo().frame("f_query_add_mod");
   					driver.switchTo().frame("framePhysicalInventoryDetail");
					break;
					
				case "framePhysicalInventoryListDetail":
                	driver.switchTo().frame("content");
                	driver.switchTo().frame("f_query_add_mod");
                	driver.switchTo().frame("framePhysicalInventoryEntryList");
   					driver.switchTo().frame("framePhysicalInventoryListDetail");
					break;
					
				case "f_button_1":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_detail");
                    driver.switchTo().frame("f_ord_detail");
                    driver.switchTo().frame("f_drug_detail");
                    driver.switchTo().frame("f_button_1");
                    break;
                    
				case "f_button_5":
                    driver.switchTo().frame("content");
                    driver.switchTo().frame("f_detail");
                    driver.switchTo().frame("f_ord_detail");
                    driver.switchTo().frame("f_drug_detail");
                    driver.switchTo().frame("f_button3");
                    break;
					
				case "Policy No":
						driver.switchTo().frame(driver.findElement(By.xpath("//frame[contains(@src,'AddModifyPatFinDetailsInsBodyEdit')]")));
					break;
					
				case "Sensitive_Diagnosis":
					driver.switchTo().frame(driver.findElement(By.xpath("//frame[contains(@src,'../../eMR/jsp/AuthorizeMRAccess.jsp')]")));
					break;	
					
				case "RecClinicalNotesTemplateFrame5":
					driver.switchTo().frame("content");
					driver.switchTo().frame("workAreaFrame");
					driver.switchTo().frame("RecClinicalNotesFrame");
					driver.switchTo().frame("RecClinicalNotesSecDetailsFrame");
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("RecClinicalNotesTemplateFrame"));
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("C_MEDCERTNUM3"));
					break;
					
				case "frame6":
                    driver.switchTo().frame("frame2");
                    break;
                    
				case "frame7":
                    driver.switchTo().frame("frame3");
                    break;

				case "result6":
					driver.switchTo().frame("content");	
					driver.switchTo().frame("f_query_add_mod");
					driver.switchTo().frame("AEMPSearchResultFrame");	
					driver.switchTo().frame("result");
					break;
					
				case "frameIssueDetail" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("frameIssueDetail"); 
		                 break;

		case "f_query_add_mod_query1" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("f_query_add_mod_query"); 
		                 break;

		case "f_SecondaryTriageCriteria" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("f_SecondaryTriageCriteria"); 
		                 break;

		case "f_SecondaryTriageResult" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("f_SecondaryTriageResult"); 
		                 break;
		case "f_attendanceCriteria" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("f_attendanceCriteria"); 
		                 break;

		case "f_attendanceResult" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("f_attendanceResult"); 
		                 break;

		case "f_patientByCriteria" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("f_patientByCriteria"); 
		                 break;

		case "f_patientByCriteria_Result" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("f_patientByCriteria_Result"); 
		                 break;
		case "frameBatchSearchQueryCriteria1" :
		           	   driver.switchTo().frame("frameBatchSearchQueryCriteria"); 
		                 break;

		case "unitFrame" :
		           	   driver.switchTo().frame("content");
				   driver.switchTo().frame("workAreaFrame");
		           	   driver.switchTo().frame("details");
		                 driver.switchTo().frame("unitFrame"); 
		                 break;
		case "ClinicalEventTypeSeqFrame" :
		           	   driver.switchTo().frame("content");
		                   driver.switchTo().frame("ClinicalEventTypeSeqFrame"); 
		                 break;

		case "IssueReturnQueryHeader" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("IssueReturnQueryHeader"); 
		                 break;

		case "IssueReturnQueryResult" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("IssueReturnQueryResult"); 
		                 break;

		case "StockTransferQueryHeader" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("StockTransferQueryHeader"); 
		                 break;

		case "StockTransferQueryResult" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("StockTransferQueryResult"); 
		                 break;

		case "frameStockTransferHeader" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("frameStockTransferHeader"); 
		                 break;

		case "SalesQueryHeader" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("SalesQueryHeader"); 
		                 break;

		case "SalesQueryResult" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("f_query_add_mod");
		                 driver.switchTo().frame("SalesQueryResult"); 
		                 break;

		case "MenstrualHistDtlQueryFrame" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("workAreaFrame");
		                 driver.switchTo().frame("MenstrualHistDtlQueryFrame"); 
		                 break;

		case "MenstrualHistDtlBottomFrame" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("workAreaFrame");
				   driver.switchTo().frame("MenstrualHistDtlResultFrame");
		                 driver.switchTo().frame("MenstrualHistDtlBottomFrame"); 
		                 break;

		case "MenstrualHistDtlResultFrame" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("workAreaFrame");
				   driver.switchTo().frame("MenstrualHistDtlResultFrame");
		                 driver.switchTo().frame("MenstrualHistDtlResultFrame"); 
		                 break;

		case "MenstrualHistDtlTopFrame" :
		           	   driver.switchTo().frame("content");
		           	   driver.switchTo().frame("workAreaFrame");
				   driver.switchTo().frame("MenstrualHistDtlResultFrame");
		                 driver.switchTo().frame("MenstrualHistDtlTopFrame"); 
		                 break;
		                 
		case "Pat_Search_Criteria_Frame1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("Pat_Search_Criteria_Frame");
		case "Pat_Search_Toolbar_Frame1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("Pat_Search_Toolbar_Frame");
		case "lookup_head1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("lookup_head");
			break;
		case "lookup_head2":  
			driver.switchTo().frame("content");
			driver.switchTo().frame("f_query_add_mod");
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("lookup_head");
			break;
		case "lookup_detail1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("lookup_detail");
			break;
		case "lookup_detail2":
			driver.switchTo().frame("content");
			driver.switchTo().frame("f_query_add_mod");
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("lookup_detail");
			break;
		case "MainFrame11":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("MainFrame1");
			break;
		case "group_head2":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("group_head");
		break;
		case "frame01":
			driver.switchTo().frame("content");
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("frame0");
			break;
		case "main1":
			driver.switchTo().frame("content");
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("main");
			break;
		case "qryCriteria1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("qryCriteria");
			break;
		case "frameBatchSearchQueryResult2":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("frameBatchSearchQueryResult");
			break;
		case "frameBatchSearchQueryCriteria2":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("frameBatchSearchQueryCriteria");
			break;
		case "frameSalesReturnSearchWindowCriteria1":
			driver.switchTo().frame("content");
			driver.switchTo().frame("f_query_add_mod");
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("frameSalesReturnSearchWindowCriteria");
			break;
		case "RIAuthorizeQueryHeaderFrame1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("RIAuthorizeQueryHeaderFrame");
			break;
		case "RIAuthorizeQueryResultFrame1":
			driver.switchTo().frame("dialog-body");
			driver.switchTo().frame("RIAuthorizeQueryResultFrame");
			break;
		 case "commontoolbar2":
			 driver.switchTo().frame("dialog-body");
				driver.switchTo().frame("commontoolbar");
				break;		
		 case "dialog-body":
			 driver.switchTo().frame("dialog-body");
				break;
		 case "dialog-body1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("f_query_add_mod");
			 driver.switchTo().frame("dialog-body");
				break;
		 case "dialog-body2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("dialog-body");
				break;
		 case "qryResult1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("qryResult");
				break;
		 case "IssueDocNoSearchHeaderFrame1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("f_query_add_mod");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("IssueDocNoSearchHeaderFrame");
				break;
		 case "IssueDocNoSearchResultFrame1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("f_query_add_mod");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("IssueDocNoSearchResultFrame");
				break;
		 case "frameBatchSearchGoodsReturnToVendorQueryResult1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("frameBatchSearchGoodsReturnToVendorQueryResult");
				break;
		 case "frameBatchSearchGoodsReturnToVendorQueryCriteria1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("frameBatchSearchGoodsReturnToVendorQueryCriteria");
				break;
		 case "mealMenu1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("mealMenu");
				break;
		 case "test1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("test");
				break;
		 case "SpecimenCollectionDateButton1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("SpecimenCollectionDateButton");
				break;
		 case "SpecimenDispatchDetail1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("SpecimenDispatchDetail");
				break;
		 case "SpecimenDispatchDetailButton1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("SpecimenDispatchDetailButton");
				break;
		 case "lookup_footer1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("lookup_footer");
				break;
		 case "lookup_header1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("lookup_header");
				break;
		 case "f_preview_buttons4":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("f_preview_buttons");
				break;
		 case "printSelectFrame1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("printSelectFrame");
				break;
		 case "buttonFrame2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("buttonFrame");
				break;
		 case "AssignRelationshipFrame2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("AssignRelationshipFrame");
				break;
		 case "chartRecordingBottomFrame1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("chartRecordingBottomFrame");
				break;
		 case "dialog-body3":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("orderDetailFrame");
			 driver.switchTo().frame("mainFrame");
			 driver.switchTo().frame("DetailFrame");
			 driver.switchTo().frame("criteriaPlaceOrderFrame");
			 driver.switchTo().frame("dialog-body");
				break;
		 case "RecPatChiefComplaintDiagLookupResultFrame2":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecPatChiefComplaintDiagLookupResultFrame");
				break;
		 case "Modify High Risk2":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("Modify High Risk");
				break;
		 case "detailpage1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("detailpage");
			 break;
		 case "viewdetailpage1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("viewdetailpage");
			 break;
		 case "RecClinicalNotesLinkHistNoteCriteriaFrame1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesLinkHistNoteCriteriaFrame");
			 break;
		 case "RecClinicalNotesLinkHistNoteTreeFrame1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesLinkHistNoteTreeFrame");
			 break;
		 case "RecClinicalNotesLinkHistNoteAddButtonsFrame1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesLinkHistNoteAddButtonsFrame");
			 break;
		 case "RecClinicalNotesLinkHistNoteSelectButtons1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesLinkHistNoteSelectButtons");
			 break;
		 case "refusal_resultframe12":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("refusal_resultframe1");
				break;
		 case "RecClinicalNotesSearchToolbarFrame2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesSearchToolbarFrame");
				break;
		 case "criteria5":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("criteria");
				break;
		 case "OrderFormat1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("OrderFormat");
				break;
		 case "flex_fields_button1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("flex_fields_button");
				break;
		 case "DetailFrame4":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("DetailFrame");
				break;
		 case "RecordButton4":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecordButton");
				break;
		 case "RecClinicalNotesMedServiceCriteriaFrame1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("RecClinicalNotesFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesMedServiceCriteriaFrame");
				break;
		 case "RecClinicalNotesMedServiceResultFrame1":
			 driver.switchTo().frame("content");
			 driver.switchTo().frame("workAreaFrame");
			 driver.switchTo().frame("RecClinicalNotesFrame");
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecClinicalNotesMedServiceResultFrame");
				break;
		 case "detail_frame5":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("detail_frame");
				break;
		 case "RecordButton_frame1":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("RecordButton_frame");
				break;
		 case "main2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("main");
				break;
		 case "MedBoardRequestFormationMain2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("MedBoardRequestFormationMain");
				break;
		 case "MedBoardRequestButtons2":
			 driver.switchTo().frame("dialog-body");
			 driver.switchTo().frame("MedBoardRequestButtons");
				break;
		 case "code_label2":
			 driver.switchTo().frame("RecDiagnosisMain");
			 driver.switchTo().frame("dialog-body");
				driver.switchTo().frame("TermCodeSearchByFrame");
				driver.switchTo().frame("code_label");
				break;		
		 case "codedesc2":
			  	driver.switchTo().frame("RecDiagnosisMain");
			  	driver.switchTo().frame("dialog-body");
				driver.switchTo().frame("TermCodeSearchByFrame");
				driver.switchTo().frame("codedesc");
				break;
				
				
				
				default:
										try{
							driver.switchTo().frame(framename);
							break;
					}
					catch(Exception e)
					{
						System.out.println("catch frame");
						driver.switchTo().frame(0);
						
					}
                     
				}
			}
		
			catch(Exception e){}
					try{	
					framename=(String) jsExecutor.executeScript("return self.name");
					System.out.println("framename....................."+framename+"\n");
					//return framename;
					}
					catch(Exception e){}
					return framename;
		}
             
	public static String window(WebDriverWait wait,int winnumber,WebDriver driver,String ProValue1) throws InterruptedException
	{
		try{
			System.out.println("**Window handling started**"+winnumber);
			wait.until(ExpectedConditions.numberOfWindowsToBe(winnumber));
			Set<String> allWindows1 = driver.getWindowHandles();
			System.out.println("Actual Window Count "+allWindows1.size());
			for(String window:allWindows1)
			{
				 String title = driver.switchTo().window(window).getTitle();
			/*	 System.out.println("title "+title);
				 System.out.println("ProValue1 "+ProValue1);*/
				 title=title.toLowerCase().trim();
				 
				 ProValue1=ProValue1.toLowerCase().trim();
				 if(title.contains(ProValue1))
				 {
					 break;
				 }
				
			 }

			
				System.out.println("Currently the driver's pointer is in the window " + driver.getTitle());
				
			}
			catch(Exception e){}
			String title=driver.getTitle();
			return title;
	
	}


	
	@SuppressWarnings("deprecation")
	public static void switchToNewWindow(WebDriverWait wait,int windowNumber,WebDriver driver,String ProValue1,String Module) throws InterruptedException 
    {	
		try{
		wait.until(ExpectedConditions.numberOfWindowsToBe(windowNumber));
           
           if(Module.equals("arraywindow"))
           {
        	   	Set<String> AllWindowHandles = driver.getWindowHandles();
           for(String win:AllWindowHandles)
           {
                  driver.switchTo().window(win);
                  String cururl=driver.getCurrentUrl();
                  if(cururl.contains(ProValue1))///pass it from test case sheet
                  {
                  System.out.println("driver navigated to array window   "+driver.getTitle());
                  break;
                  }
           }
           }
           else if(Module.equals("arraywindowwinclose"))
           {

                  
                  Set<String> AllWindowHandles = driver.getWindowHandles();
           for(String win:AllWindowHandles)
           {
                  driver.switchTo().window(win);
                  String cururl=driver.getCurrentUrl();
                  if(cururl.contains(ProValue1))///pass it from test case sheet
                  {
                  driver.close();
                  break;
                  }
           }
           
           }
           else
           {
           Set < String > s = driver.getWindowHandles();   
        Iterator < String > ite = s.iterator();
        int i = 1;
        while (ite.hasNext() && i < windowNumber+1) {
            String popupHandle = ite.next().toString();
          driver.switchTo().window(popupHandle);
          String title=driver.getTitle();
            title=title.toLowerCase().trim();
            ProValue1=ProValue1.toLowerCase().trim();
            if(ProValue1.contains(" -- webpage dialog"))
				{
					String ProValue1splt[] = ProValue1.split(" -- webpage dialog");
					ProValue1=ProValue1splt[0];
				}
            
            if (title.startsWith(ProValue1)) 
            {
            	 System.out.println("Currently the driver's pointer switched to new window : "+driver.getTitle()+"\n");
                  break;
            }
                  
            i++;
    }
           }}
		catch(Exception e){}
           }

	

	
	public static void currentwindow(WebDriver driver,String ProValue1,WebDriverWait wait, int winnumber,String Module) throws InterruptedException 
	{
		
		String Browserdec=driver.switchTo().window(driver.getWindowHandle()).getTitle();
			if(Browserdec.startsWith(ProValue1))
			{
				System.out.println("Currently the driver's pointer is in the window "+driver.getTitle());
			}
		
			else
			{
				Window_Frame_Handling.switchToNewWindow(wait,winnumber,driver,ProValue1,Module);
				
			}
		
	}
	public static void selectOption(WebDriver wdriver,WebElement wele,String value) throws Exception {
	      JavascriptExecutor jexec = (JavascriptExecutor) wdriver;
	      // jexec.executeScript("var select = arguments[0];  for(var i = 0; i < select.options.length; i++){ if(select.options[i].text == arguments[1]){ select.options[i].selected = true; } } ", wele, value);
	      // jexec.executeScript("var select = arguments[0];  for(var i = 0; i < select.options.length; i++){ if(select.options[i].text == arguments[1]){ select.options[i].selected = true; } } var evt = document.createEvent('HTMLEvents'); evt.initEvent('change', true, true); select.dispatchEvent(evt);", wele, value);
	      jexec.executeScript("var select = arguments[0];  for(var i = 0; i < select.options.length; i++){ if(select.options[i].text == arguments[1]){ select.options[i].selected = true; } } if ('createEvent' in document) { var evt = document.createEvent('HTMLEvents'); evt.initEvent('change', true, true); select.dispatchEvent(evt); } else { select.fireEvent('onchange'); }", wele, value);
	      // jexec.executeScript("var select = arguments[0]; if ('createEvent' in document) { var evt = document.createEvent('HTMLEvents'); evt.initEvent('change', false, true); element.dispatchEvent(evt); } else { element.fireEvent('onchange'); }  for(var i = 0; i < select.options.length; i++){ if(select.options[i].text == arguments[1]){ select.options[i].selected = true; } } ", wele, value);
	  }
	
}
