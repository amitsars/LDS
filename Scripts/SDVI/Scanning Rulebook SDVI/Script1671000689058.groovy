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
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.By as By
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

WebUI.callTestCase(findTestCase('CONV/AdUPN'), [:], FailureHandling.STOP_ON_FAILURE)

//reading excel data

FileInputStream file = new FileInputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))
XSSFWorkbook workbook = new XSSFWorkbook(file)
XSSFSheet sheet = workbook.getSheet('SDVI')
XSSFSheet sheet1 = workbook.getSheet('TC_Status')
XSSFSheet sheet2 = workbook.getSheet('CONV')


String environment = sheet.getRow(1).getCell(1).getStringCellValue();
String environment_dev = sheet.getRow(1).getCell(4).getStringCellValue();
String environment_qa = sheet.getRow(2).getCell(4).getStringCellValue();
String SG_Input1 = sheet.getRow(2).getCell(1).getStringCellValue();
String SG_Input2 = sheet.getRow(3).getCell(1).getStringCellValue();
String SG_Input3 = sheet.getRow(4).getCell(1).getStringCellValue();
//String SG_Input4 = sheet.getRow(5).getCell(1).getStringCellValue();
String SG_Input5 = sheet.getRow(6).getCell(1).getStringCellValue();


if(environment.contains("Dev")) {
WebUI.navigateToUrl(environment_dev)
}
else {
	WebUI.navigateToUrl(environment_qa)
}
	
WebUI.mouseOver(findTestObject('Object Repository/VA-SDVI/Page_View Rulebooks  Underwriting/a_Rulebooks'))

WebUI.click(findTestObject('Object Repository/VA-SDVI/Page_View Rulebooks  Underwriting/a_View Rulebooks'))
//driver.sleep(1000)
WebUI.click(findTestObject('VA-SDVI/Page_View Rulebooks  Underwriting/a_Edit'))

WebUI.click(findTestObject('Object Repository/VA-SDVI/Page_ScanningSwRulebook  Edit Rulebook  Und_905085/a_Start a new test interview using this rulebook'))

WebUI.setText(findTestObject('Object Repository/Scanning/Page_Start Interview  Underwriting/input_Document ID_interview.readOnlyScores0_ebbbaa'), 
    SG_Input1)

WebUI.setText(findTestObject('Object Repository/Scanning/Page_Start Interview  Underwriting/input_Document System of Record - sql,onbas_c8c9a2'), 
    SG_Input2)

WebUI.setText(findTestObject('Object Repository/Scanning/Page_Start Interview  Underwriting/input_SG_INPUT3_interview.readOnlyScores2.t_99f475'), 
    SG_Input3)

WebUI.setText(findTestObject('Object Repository/Scanning/Page_Start Interview  Underwriting/input_SG_INPUT5_interview.readOnlyScores4.t_0f582a'), 
    SG_Input5)

WebUI.click(findTestObject('Object Repository/Scanning/Page_Start Interview  Underwriting/input_Nothing found to display_methodexecuteOnce'))

WebUI.delay(5)
try {
def NewSDVIPolicy = WebUI.getText(findTestObject('Object Repository/Scanning/GetinterviewDetails/Get Policy'))
	sheet.getRow(8).createCell(1).setCellValue(NewSDVIPolicy);
	sheet2.getRow(5).createCell(1).setCellValue(NewSDVIPolicy);
	sheet1.getRow(1).createCell(3).setCellValue("Pass");
	sheet1.getRow(2).createCell(3).setCellValue("Pass");
}catch (Exception E)
	{
		sheet.getRow(8).createCell(5).setCellValue("No Policy was created as existing SSN was used");
		sheet1.getRow(2).createCell(3).setCellValue("Fail");
		file.close()
		FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))
		workbook.write(outFile)
		outFile.close()
	}
	
def ExistingWI = WebUI.getText(findTestObject('Object Repository/Scanning/GetinterviewDetails/GetWorkitemID'))
sheet.getRow(8).createCell(3).setCellValue(ExistingWI);
	
	
file.close()
FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))
workbook.write(outFile)
outFile.close()
	
WebUI.delay(5)
//WebUI.closeBrowser()


