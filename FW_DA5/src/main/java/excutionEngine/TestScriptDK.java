package excutionEngine;
import java.lang.reflect.Method;

import org.junit.Test;

import utilities.ActionKeywords;
import utilities.ExcelUtils;
import org.junit.runner.JUnitCore;
import org.junit.runner.Result;
import org.junit.runner.notification.Failure;

public class TestScriptDK {

	@Test
	public void excute_TestCasedk() throws Exception {

		String sPath = System.getProperty("user.dir") + "//src//main//java//data_Engine//Data_ÐA5.xlsx";
		ExcelUtils.setExcelFile(sPath, "ÐK");
		int CasePass=0;
        int CaseFail=0;
        int CaseSkip=0;
        int row = ExcelUtils.getRowCount("ÐK");
		for (int i = 1; i <= 88; i++) {
			
				System.out.println("Line:"+ i);
				String sActionKeyword = ExcelUtils.getCellData(i, 3);
				String locatorType = ExcelUtils.getCellData(i, 5);

				String locatorValue = ExcelUtils.getCellData(i, 6);
				String testData = ExcelUtils.getCellData(i, 7);

				switch (sActionKeyword) {
				case "openBrowser":
					ActionKeywords.openBrowser(testData);
					break;
				case "navigate":
					ActionKeywords.navigate(testData);
					break;
				case "setText":
					ActionKeywords.setText(locatorType, locatorValue, testData);
					break;
				case "click":
					ActionKeywords.clickElement(locatorType, locatorValue);
					break;
				case "verifyText2":
					if (ActionKeywords.verifyText2(locatorType, locatorValue, testData)) {
                   	 System.out.println("Same result ---> pass");
                        CasePass++;
                        // ExtentTestManager.logMessage(Status.PASS,testData);
                    } else {
                   	 System.out.println("Different result ---> Fail");
                        CaseFail++;
                        // ExtentTestManager.logMessage(Status.FAIL,testData);
                    }
                    break;
				case "quitBrowser":
					ActionKeywords.quitDriver();
					break;
				default:
					System.out.println("[>>ERROR<<]: |Keyword Not Found " + sActionKeyword);
				}
			}
		java.util.Date date=new java.util.Date();
        System.out.println("==========================================================");
        System.out.println("-----------"+date+"--------------");
        System.out.println("Total number of Testcases run: "+(CasePass+CaseFail+CaseSkip));
        System.out.println("Total number of passed Testcases: "+CasePass);
        System.out.println("Total number of failed Testcases: "+CaseFail);
        System.out.println("Total number of skip Testcases: "+CaseSkip);
        System.out.println("==========================================================");
		}
	}
