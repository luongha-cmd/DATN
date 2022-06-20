package utilities;
import static org.junit.Assert.assertTrue;

import java.util.List;
import java.util.NoSuchElementException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
//import org.apache.log4j.Logger;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
//import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.opera.OperaDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;
public class ActionKeywords {

//		public static Logger APP_LOGS = Logger.getLogger("devpinoyLogger");

		public static WebDriver driver;
		private static final int timeoutWait = 10;
		private static final int timeoutWaitForPageLoaded = 20;
		private static Actions action;
		private static JavascriptExecutor js;
		private static WebDriverWait wait;

//		public ActionKeywords() {
//			System.setProperty("org.apache.commons.logging.Log", "org.apache.commons.logging.impl.Jdk14Logger");
//		}

		private static WebElement GetElement(String locatorType, String locatorValue) {
			WebElement element = null;

			if (locatorType.equalsIgnoreCase("className"))
				element = driver.findElement(By.className(locatorValue));
			else if (locatorType.equalsIgnoreCase("cssSelector"))
				element = driver.findElement(By.cssSelector(locatorValue));
			else if (locatorType.equalsIgnoreCase("id"))
				element = driver.findElement(By.id(locatorValue));
			else if (locatorType.equalsIgnoreCase("partialLinkText"))
				element = driver.findElement(By.partialLinkText(locatorValue));
			else if (locatorType.equalsIgnoreCase("name"))
				element = driver.findElement(By.name(locatorValue));
			else if (locatorType.equalsIgnoreCase("xpath"))
				element = driver.findElement(By.xpath(locatorValue));
			else if (locatorType.equalsIgnoreCase("tagName"))
				element = driver.findElement(By.tagName(locatorValue));
			else {
				System.out.println("[>>ERROR<<]: |GetElement |" + locatorType + "=" + locatorValue);
//			  System.out.println("Format locator suppoted:\n *id\n *name\n *css\n *link\n *xpath\n"); 
			}

			return element;
		}

		public By getBy(String locatorType, String locatorValue) {
			By by = null;

			if (locatorType.toLowerCase().contains("id")) {
				by = By.id(locatorValue);
			} else if (locatorType.toLowerCase().contains("name")) {
				by = By.name(locatorValue);
			} else if (locatorType.toLowerCase().contains("tagname")) {
				by = By.tagName(locatorValue);
			} else if (locatorType.toLowerCase().contains("classname")) {
				by = By.className(locatorValue);
			} else if (locatorType.toLowerCase().contains("cssselector")) {
				by = By.cssSelector(locatorValue);
			} else if (locatorType.toLowerCase().contains("linktext")) {
				by = By.linkText(locatorValue);
			} else if (locatorType.toLowerCase().contains("partiallinktext")) {
				by = By.partialLinkText(locatorValue);
			} else if (locatorType.toLowerCase().contains("xpath")) {
				by = By.xpath(locatorValue);
			} else {
				System.out.println("[>>ERROR<<]: |getBy |" + locatorType + "=" + locatorValue);
//			  System.out.println("Format locator suppoted:\n *id\n *name\n *css\n *link\n *xpath\n"); 
			}
			return by;
		}
		//cuộn chuột
		public static void hover()
		{
			JavascriptExecutor js = (JavascriptExecutor) driver;// cuộn chuột
		    js.executeScript("window.scrollBy(0,700)");
		}   
		
		public static void hover2()
		{
			JavascriptExecutor js = (JavascriptExecutor) driver;// cuộn chuột
		    js.executeScript("window.scrollBy(0,500)");
		}   
		//di chuột
		public static void ElementPerfom(String address)
		{
			try 
			{
				if(address.contains("["))
				{
					Actions a = new Actions(driver);
				      a.moveToElement(driver.findElement(By.xpath(address))).
				     // contextClick().
				      build().perform();
				}else 
				{
					Actions a = new Actions(driver);
				      a.moveToElement(driver.findElement(By.id(address))).build().perform();
				}
			}catch(Exception e) {
				System.out.println("===============Khong tim thay dia chi:");
			}
		}
		
//		String actuakResult1=driver.switchTo().alert().getText();
//		assertTrue (actuakResult1.contains("Đăng ký thành công!"));
		public static void alert()
		{
//				String actuakResult1=driver.switchTo().alert().getText();
//				assertTrue (actuakResult1.contains("Đăng ký thành công!"));
				String actuakResult1=driver.switchTo().alert().getText();
				if (actuakResult1.contains("Đăng ký thành công!")==true)
					driver.switchTo().alert().accept();
		}                        
		public static WebDriver openBrowser(String browserName) {
//			APP_LOGS.debug("[info] Executing: |open browser: "+browserName);
			System.out.println("[info] Executing: |open browser: " + browserName);
//			System.out.println();
			switch (browserName.trim().toLowerCase()) {
			case "chrome":
				driver = initChromeDriver();
				break;
			case "firefox":
				driver = initFirefoxDriver();
				break;
			case "opera":
				driver = initOperaDriver();
				break;
			default:
				System.out.println("[>>ERROR<<]: |open browser: " + browserName);
				// System.out.println("Browser: "+browserName+" is invalid, Launching Chrome as
				// browser of choice...");
				driver = initChromeDriver();
			}
			wait = new WebDriverWait(driver, timeoutWait);
			return driver;
		}

		public static void navigate(String appURL) {
			try {
//				APP_LOGS.debug("[info] Executing: |open Url:| " + appURL + "|application|");
				System.out.println("[info] Executing: |open Url:| " + appURL + "|application|");
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.navigate().to(appURL);
			} catch (Exception e) {
				System.out.println("[info] Executing: |open Url:| " + appURL + "|application|");
				// System.out.println("Error..."+e.getStackTrace());
			}

		}

		// Ham nay tuy chon Browser. Cho chay truoc khi goi class nay (BeforeClass)
		private static void setDriver(String browserName, String appURL) {
			switch (browserName) {
			case "chrome":
				driver = initChromeDriver();
				driver.navigate().to(appURL);
				break;
			case "firefox":
				driver = initFirefoxDriver();
				driver.navigate().to(appURL);
				break;
			case "opera":
				driver = initOperaDriver();
				driver.navigate().to(appURL);
				break;
			default:
				System.out.println("[info] Executing: |open Url:| " + browserName
						+ "|is invalid, Launching Chrome as browser of choice...|");
				System.out.println("Browser: " + browserName + " is invalid, Launching Chrome as browser of choice...");
				driver = initChromeDriver();
			}
		}

//		//Ham khoi tao trinh duyet va duong dan ung dung
//		public static void openBrowser(String browserType, String appURL) {
//			try {
//				//Khoi tao driver va tuy chon browser vÃ  web url
//				setDriver(browserType,appURL);
//			}catch(Exception e) {
//				System.out.println("Error..."+e.getStackTrace());
//			}
//		}

		public static void quitDriver() throws Exception {
			if (driver.toString().contains("null")) {
				System.out.println("[info] Executing: All Browser windows are closed");
//				System.out.print("All Browser windows are closed ");
			} else {
				System.out.println("[info] Executing: |quit|");
				driver.manage().deleteAllCookies();
				Thread.sleep(2000);
				driver.quit();
			}
		}

		public static void setText(WebElement element, String value) {
			try {
				System.out.println("[info] Executing: |setText| element: " + element + ", sendKey: " + value);
//				APP_LOGS.debug("[info] Executing: |setText| element: " + element +", sendKey: " + value);
				waitForPageLoaded();
				wait.until(ExpectedConditions.visibilityOf(element));
				element.clear();
				element.sendKeys(value);
			} catch (NoSuchElementException e) {
				System.out.println(
						"[>>ERROR<<]: |setText| element: " + element + " not found to sendKeys| " + e.getMessage());
			}

		}

		public static void setText(String locatorType, String locatorValue, String value) {
			try {
				System.out.println("[info] Executing: |sendKeys| element: " + locatorType + "= " + locatorValue
						+ ", sendKey: " + value);
//				APP_LOGS.debug("[info] Executing: |sendKeys| element: " + locatorType + "= "+ locatorValue + ", sendKey: " + value);
				WebElement element = GetElement(locatorType, locatorValue);
				waitForPageLoaded();
				wait.until(ExpectedConditions.visibilityOf(element));
				element.clear();
				element.sendKeys(value);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |sendKeys: " + locatorType + "= " + locatorValue
						+ " not found to sendKeys| " + e.getMessage());
			}
		}

		// Khoi tao cau hinh cua cac Browser de dua vao Switch Case
		private static WebDriver initChromeDriver() {
			System.out.println("Launching Chrome browser...");
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			return driver;
		}

		private static WebDriver initFirefoxDriver() {
			System.out.println("Launching Firefox browser...");
			WebDriverManager.firefoxdriver().setup();
			driver = new FirefoxDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			return driver;
		}

		private static WebDriver initOperaDriver() {
			System.out.println("Launching Opera browser...");
			WebDriverManager.operadriver().setup();
			driver = new OperaDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			return driver;
		}

		public static void clickElement(WebElement element) {
			try {
				System.out.println("[info] Executing: |click| element: " + element);
//				APP_LOGS.debug("[info] Executing: |click| element: " + element );
				waitForPageLoaded();
				wait.until(ExpectedConditions.visibilityOf(element));
				element.click();
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |click: " + element + " not found to click| " + e.getMessage());
			}
		}

		public static void clickElement(String locatorType, String locatorValue) {
			try {
				System.out.println("[info] Executing: |click| element: " + locatorType + "= " + locatorValue);
//				APP_LOGS.debug("[info] Executing: |click| element: " + locatorType + "= "+ locatorValue);
				WebElement element = GetElement(locatorType, locatorValue);
				waitForPageLoaded();
				wait.until(ExpectedConditions.visibilityOf(element));
				element.click();
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |click: " + locatorType + "= " + locatorValue + " not found to click| "
						+ e.getMessage());
			}
		}

//		public static void clickElement(String element) {
//			waitForPageLoaded();
////			wait.until(ExpectedConditions.visibilityOf(element));
//			driver.findElement(By.xpath(OR.getProperty(element))).click();
//		}

		public static void clickElementWithJs(WebElement element) {
			try {
				System.out.println("[info] Executing: |clickElementWithJs| element: " + element);
				js.executeScript("arguments[0].scrollIntoView(true)", element);
				wait.until(ExpectedConditions.elementToBeClickable(element));
				js.executeScript("arguments[0].click();", element);
			} catch (NoSuchElementException e) {
				System.out
						.println("[>>ERROR<<]: |clickElementWithJs: " + element + " not found to click| " + e.getMessage());
			}
		}

		public static void clickElementWithJs(String locatorType, String locatorValue) {
			try {
				System.out.println("[info] Executing: |clickElementWithJs| element: " + locatorType + "= " + locatorValue);
				WebElement element = GetElement(locatorType, locatorValue);
				js.executeScript("arguments[0].scrollIntoView(true)", element);
				wait.until(ExpectedConditions.elementToBeClickable(element));
				js.executeScript("arguments[0].click();", element);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |clickElementWithJs: " + locatorType + "= " + locatorValue
						+ " not found to click| " + e.getMessage());
			}
		}

		public static void waitForPageLoaded() {
			try {
				System.out.println("[info] Executing: |Timeout waiting for Page Load request");
				wait.until(new Function<WebDriver, Boolean>() {
					public Boolean apply(WebDriver driver) {
						return String.valueOf(((JavascriptExecutor) driver).executeScript("return document.readyState"))
								.equals("complete");
					}
				});
			} catch (Throwable e) {
				System.out.println("[>>ERROR<<]: |Timeout waiting for Page Load request" + e.getMessage());
//				System.out.println("Timeout waiting for Page Load request.");			
			}
		}

		public static boolean verifyPageLoaded(String pageLoadedText) {
			System.out.println("[info] Executing: |verify page loaded");
			waitForPageLoaded();
			Boolean res = false;

			List<WebElement> elementList = driver.findElements(By.xpath("//*contains(text(),'" + pageLoadedText + "')]"));
			if (elementList.size() > 0) {
				res = true;
				System.out.println("[info] Executing: |Page loaded(" + res + "): " + pageLoadedText);
//				System.out.println("Page loaded(" + res + "): " + pageLoadedText);
			} else {
				res = false;
				System.out.println("[info] Executing: |Page loaded(" + res + "): " + pageLoadedText);
//				System.out.println("Page loaded(" + res + "): " + pageLoadedText);
			}
			return res;
		}

		public static boolean verifyUrl(String url) {
//			System.out.println(driver.getCurrentUrl());
//			System.out.println(url);
			System.out.println("[info] Executing: |Verify Url");
			return driver.getCurrentUrl().contains(url);
		}

		public static String getPageTitle() {
			System.out.println("[info] Executing: |Get Page Title");
			waitForPageLoaded();
			return driver.getTitle();
		}

		public static boolean verifyPageTitle(String pageTitle) {
			System.out.println("[info] Executing: |Verify Page Title");
			return getPageTitle().equals(pageTitle);
		}

		public static void rightClickElement(WebElement element) {
			try {
				System.out.println("[info] Executing: |rightClick| element: " + element);
				wait.until(ExpectedConditions.elementToBeClickable(element));
				action.contextClick().build().perform();
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |rightClick: " + element + " + element");
			}
		}

		public static void rightClickElement(String locatorType, String locatorValue) {
			try {
				System.out.println("[info] Executing: |Right click| element: " + locatorType + "= " + locatorValue);
				WebElement element = GetElement(locatorType, locatorValue);

				wait.until(ExpectedConditions.elementToBeClickable(element));
				action.contextClick().build().perform();
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |Right click: " + locatorType + "= " + locatorValue
						+ " not found to click| " + e.getMessage());
			}
		}

		// Chon du lieu tu combobox
		public static void selectOptionByText(WebElement element, String text) {
			try {
				System.out.println("[info] Executing: |select Option By Text| element: " + element + "|text: " + text);
				Select select = new Select(element);
				select.selectByVisibleText(text);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |select Option By Text| element: " + element + "|text: " + text
						+ " not found to select| " + e.getMessage());
			}
		}
		public static void selectOptionByText(String locatorType, String locatorValue, String text) {
			WebElement element =  GetElement(locatorType, locatorValue);
			try {
				System.out.println("[info] Executing: |select Option By Text| element: " + element + "|text: " + text);
				Select select = new Select(element);
				select.selectByVisibleText(text);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |select Option By Text| element: " + element + "|text: " + text
						+ " not found to select| " + e.getMessage());
			}
		}


		public static void selectOptionByValue(WebElement element, String value) {
			try {
				System.out.println("[info] Executing: |select Option By Value| element: " + element + "|value: " + value);
				Select select = new Select(element);
				select.selectByValue(value);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |select Option By Value| element: " + element + "|value: " + value
						+ " not found to select| " + e.getMessage());
			}
		}

		public static void selectOptionByValue(String locatorType, String locatorValue, String value) {
			try {
				System.out.println("[info] Executing: |select Option By Value| " + locatorType + "= " + locatorValue
						+ "|value: " + value);
				WebElement element = GetElement(locatorType, locatorValue);
				Select select = new Select(element);
				select.selectByValue(value);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |select Option By Value|: " + locatorType + "= " + locatorValue
						+ "|value: " + value + " not found to select| " + e.getMessage());
			}
		}

		public static void selectOptionByIndex(WebElement element, int index) {
			try {
				System.out.println("[info] Executing: |select Option By Index| element: " + element + "|index: " + index);
				Select select = new Select(element);
				select.selectByIndex(index);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |select Option By Index| element: " + element + "|index: " + index
						+ " not found to select| " + e.getMessage());
			}
		}

		public static void selectOptionByIndex(String locatorType, String locatorValue, int index) {
			try {
				System.out.println("[info] Executing: |select Option By Index| " + locatorType + "= " + locatorValue
						+ "|index: " + index);
				WebElement element = GetElement(locatorType, locatorValue);
				Select select = new Select(element);
				select.selectByIndex(index);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |select Option By Index|: " + locatorType + "= " + locatorValue
						+ "|index: " + index + " not found to select| " + e.getMessage());
			}
		}

		public static boolean verifyElementText(WebElement element, String textValue) {
			boolean result = true;
			try {
				System.out.println("[info] Executing: |verify Element Text| element: " + element + "|Value: " + textValue);
//				APP_LOGS.debug("[info] Executing: |verify Element Text| element: " + element + "|Value: " + textValue);
				wait.until(ExpectedConditions.visibilityOf(element));
				result = element.getText().equals(textValue);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |verify Element Text| element: " + element + "|Value: " + textValue
						+ " not found to verify| " + e.getMessage());
//				APP_LOGS.debug("[>>ERROR<<]: |verify Element Text| element: " + element + "|Value: " + textValue +" not found to verify| " + e.getMessage());
				result = false;
			}
			return result;
		}

		public static boolean verifyElementText(String locatorType, String locatorValue, String text) {
			boolean result = true;
			try {
				System.out.println("[info] Executing: |verify Element Text| element: " + locatorType + "= " + locatorValue
						+ "|text: " + text);
//				APP_LOGS.debug("[info] Executing: |verify Element Text| element: " + locatorType + "= "+ locatorValue + "|text: " + text);
				WebElement element = GetElement(locatorType, locatorValue);
				wait.until(ExpectedConditions.visibilityOf(element));
				result = element.getText().equals(text);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |verify Element Text| element: " + locatorType + "= " + locatorValue
						+ "|text: " + text + " not found to verify| " + e.getMessage());
//				APP_LOGS.debug("[>>ERROR<<]: |verify Element Text| element: " + locatorType + "= "+ locatorValue + "|text: " + text + " not found to verify|" + e.getMessage());
				result = false;
			}
			return result;
		}
	//
		public static boolean verifyElementValue(WebElement element, String value) {
			wait.until(ExpectedConditions.visibilityOf(element));
			return element.getAttribute("value").equals(value);
		}

		public static boolean verifyElementValue(String locatorType, String locatorValue, String value) {
			WebElement element = GetElement(locatorType, locatorValue);
			wait.until(ExpectedConditions.visibilityOf(element));
			return element.getAttribute("value").equals(value);
		}

		public static boolean verifyElementExist(By elementBy) {
			// Táº¡o list lÆ°u táº¥t cáº£ cÃ¡c Ä‘á»‘i tÆ°á»£ng WebElement
			List<WebElement> listElement = driver.findElements(elementBy);

			int total = listElement.size();
			if (total > 0) {
				return true;
			}
			return false;
		}

		// wait for Javascript to Loaded
		public static void waitForJQueryLoaded() {
			ExpectedCondition<Boolean> jQueryload = new ExpectedCondition<Boolean>() {
				public Boolean apply(WebDriver driver) {
					try {

						return ((Long) ((JavascriptExecutor) driver).executeScript("return jQuery.active") == 0);
					} catch (Exception e) {
						return true;
					}
				}
			};
			try {
				wait = new WebDriverWait(driver, timeoutWaitForPageLoaded);
				wait.until(jQueryload);
			} catch (Throwable error) {
				System.out.println("Timeout waiting for JQuery request.");
			}
		}

		// wait for Javascript to Loaded
		public static void waitForJSLoaded() {

			ExpectedCondition<Boolean> jsload = new ExpectedCondition<Boolean>() {
				public Boolean apply(WebDriver driver) {
					return ((JavascriptExecutor) driver).executeScript("return document.readyState").toString()
							.equals("complete");
				}
			};
			try {
				wait = new WebDriverWait(driver, timeoutWaitForPageLoaded);
				wait.until(jsload);
			} catch (Throwable error) {
				System.out.println("Timeout waiting for Javascript request.");
			}
		}
		public static boolean verifyText(WebElement element, String textValue) {
			boolean result = true;
			try {
				System.out.println("[info] Executing: |verify Element Text| element: " + element + "|Value: " + textValue);
//				APP_LOGS.debug("[info] Executing: |verify Element Text| element: " + element + "|Value: " + textValue);
				wait.until(ExpectedConditions.visibilityOf(element));
				result = element.getText().equals(textValue);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |verify Element Text| element: " + element + "|Value: " + textValue
						+ " not found to verify| " + e.getMessage());
//				APP_LOGS.debug("[>>ERROR<<]: |verify Element Text| element: " + element + "|Value: " + textValue +" not found to verify| " + e.getMessage());
				result = false;
			}
			return result;
		}

		public static boolean verifyText(String locatorType, String locatorValue, String text) {
			boolean result = true;
			try {
				System.out.println("[info] Executing: |verify Element Text| element: " + locatorType + "= " + locatorValue
						+ "|text: " + text);
//				APP_LOGS.debug("[info] Executing: |verify Element Text| element: " + locatorType + "= "+ locatorValue + "|text: " + text);
				WebElement element = GetElement(locatorType, locatorValue);
				wait.until(ExpectedConditions.visibilityOf(element));
				result = element.getText().equals(text);
			} catch (NoSuchElementException e) {
				System.out.println("[>>ERROR<<]: |verify Element Text| element: " + locatorType + "= " + locatorValue
						+ "|text: " + text + " not found to verify| " + e.getMessage());
//				APP_LOGS.debug("[>>ERROR<<]: |verify Element Text| element: " + locatorType + "= "+ locatorValue + "|text: " + text + " not found to verify|" + e.getMessage());
				result = false;
			}
			return result;
		}	
		
		 public static boolean verifyText2(String locatorType, String locatorValue, String text) {
			 System.out.println("Executing: Verify Text");
		        String result;
		        WebElement element= GetElement(locatorType, locatorValue);
		        result = element.getText();
		        System.out.println("Text result: "+result);
		        System.out.println("Text expected: "+text);
		        if (text.equals(result))
		        {
		            return true;
		        }
		        else return false;
		        
		    }
}
