import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable
import org.openqa.selenium.By
import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.chrome.ChromeDriver as ChromeDriver
import org.openqa.selenium.chrome.ChromeOptions as ChromeOptions
import org.openqa.selenium.remote.DesiredCapabilities as DesiredCapabilities
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import java.util.concurrent.TimeUnit as TimeUnit

// for excel read
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
	
System.setProperty('webdriver.chrome.driver', 'C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Katalon_Studio_PE_Windows_64-8.5.0/configuration/resources/drivers/chromedriver_win32/chromedriver.exe')

ChromeOptions options = new ChromeOptions()

options.addExtensions(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/idgpnmonknjnojddfkpgkljpfnnfcklj/4.0.21_2.crx'))

DesiredCapabilities caps = new DesiredCapabilities()

caps.setCapability(ChromeOptions.CAPABILITY, options)

WebDriver driver = new ChromeDriver(caps)

DriverFactory.changeWebDriver(driver)

//reading excel data

FileInputStream file = new FileInputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))
XSSFWorkbook workbook = new XSSFWorkbook(file)
XSSFSheet sheet = workbook.getSheet('SRH')

//Read data from excel row
String AdUpnEmail = sheet.getRow(0).getCell(1).getStringCellValue();

driver.sleep(2000)
driver.get("chrome-extension://idgpnmonknjnojddfkpgkljpfnnfcklj/popup.html")
driver.sleep(2000)
WebUI.switchToWindowTitle('ModHeader')
def elem = driver.findElement(By.xpath("//input[@placeholder='Name']"))
elem.sendKeys("adUpn");
def elem1 = driver.findElement(By.xpath("//input[@placeholder='Value']"))
//elem1.sendKeys("amit.sarswat@exldemo.com");
elem1.sendKeys(AdUpnEmail);
driver.sleep(1000)

//closing excel file
file.close()

//FileOutputStream outFile = new FileOutputStream(new File('C:\\Users\\kavindra.EXLDEMO\\Desktop\\Katalon_Studio_Windows_64-8.5.2\\Excel\\TestDataUtil.xlsx'))

//workbook.write(outFile)

//outFile.close()
//close browser
//driver.close()








