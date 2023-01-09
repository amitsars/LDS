import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import java.sql.Driver as Driver
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
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.chrome.ChromeDriver as ChromeDriver
import org.openqa.selenium.chrome.ChromeOptions as ChromeOptions
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory

//WebDriver driver = new ChromeDriver(caps)
//DriverFactory.changeWebDriver(driver)

WebUI.callTestCase(findTestCase('CONV/Conv Work Item'), [:], FailureHandling.STOP_ON_FAILURE)

//WebUI.navigateToUrl('http://ec04-vc02-web01:9303/equotes/nextPage!input.action?interviewId=505605&interviewToken=wFiyqjI1k5MOpzQXhvGp&requestedLifeNumber=1&requestedPageNumber=3&requestedPageCode=&theme=exl')
//WebUI.maximizeWindow()
WebUI.delay(10)



FileInputStream file = new FileInputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheet('CONV')
XSSFSheet sheet1 = workbook.getSheet('TC_Status')


'Read data from excel'

String Postmark_Date = sheet.getRow(9).getCell(1).getStringCellValue()
String Amount_Converted = sheet.getRow(10).getCell(1).getStringCellValue()
String Amount_Continued = sheet.getRow(11).getCell(1).getStringCellValue()
String Amount_Cancelled = sheet.getRow(12).getCell(1).getStringCellValue()
String Plan_of_Insurance = sheet.getRow(13).getCell(1).getStringCellValue()
String Effective_Date = sheet.getRow(14).getCell(1).getStringCellValue()
String Payment_Method = sheet.getRow(15).getCell(1).getStringCellValue()

//String Amount_Received = sheet.getRow(42).getCell(1).getStringCellValue()
//String PB1_FirstName = sheet.getRow(45).getCell(1).getStringCellValue()
//String PB1_MiddleName = sheet.getRow(46).getCell(1).getStringCellValue()
//String PB1_LastName = sheet.getRow(47).getCell(1).getStringCellValue()
//String PB1_Relationship = sheet.getRow(48).getCell(1).getStringCellValue()
//String PB1_AddressLine1 = sheet.getRow(49).getCell(1).getStringCellValue()
//String PB1_AddressLine2 = sheet.getRow(50).getCell(1).getStringCellValue()
//String PB1_AddressLine3 = sheet.getRow(51).getCell(1).getStringCellValue()
//String PB1_City = sheet.getRow(52).getCell(1).getStringCellValue()
//String PB1_ZipCode = sheet.getRow(53).getCell(1).getStringCellValue()
//String PB1_Email = sheet.getRow(54).getCell(1).getStringCellValue()
//String PB1_State = sheet.getRow(55).getCell(1).getStringCellValue()
//String CB1_FirstName = sheet.getRow(57).getCell(1).getStringCellValue()
//String CB1_MiddleName = sheet.getRow(58).getCell(1).getStringCellValue()
//String CB1_LastName = sheet.getRow(59).getCell(1).getStringCellValue()
//String CB1_Relationship = sheet.getRow(60).getCell(1).getStringCellValue()
//String CB1_AddressLine1 = sheet.getRow(61).getCell(1).getStringCellValue()
//String CB1_AddressLine2 = sheet.getRow(62).getCell(1).getStringCellValue()
//String CB1_AddressLine3 = sheet.getRow(63).getCell(1).getStringCellValue()
//String CB1_City = sheet.getRow(64).getCell(1).getStringCellValue()
//String CB1_ZipCode = sheet.getRow(65).getCell(1).getStringCellValue()
//String CB1_Email = sheet.getRow(66).getCell(1).getStringCellValue()
//String CB1_State = sheet.getRow(67).getCell(1).getStringCellValue()

WebUI.switchToWindowTitle('eApp')
//Application Information
WebUI.click(findTestObject('Object Repository/Equotes CONV/APPLICATION INFORMATION/Continue App Information'))

WebUI.delay(5)

//Coverage Details
WebUI.setText(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Postmark Date'),
	Postmark_Date)
//WebUI.sendKeys('Object Repository/Equotes CONV/COVERAGE DETAILS/Postmark Date', Keys.chord(Keys.ENTER))
WebUI.click(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Page Header'))
WebUI.setText(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Amount Converted'),
	Amount_Converted)

WebUI.setText(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Amount Continued as Term'),
	Amount_Continued)

WebUI.setText(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Amount Cancelled'),
	Amount_Cancelled)

WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Insurance Plan'),
	Plan_of_Insurance, true)

WebUI.click(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Continue Coverage Details'))
WebUI.delay(5)
WebUI.setText(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Effective Date'),
	Effective_Date)
//WebUI.sendKeys('Object Repository/Equotes CONV/COVERAGE DETAILS/Effective Date', Keys.ENTER)
WebUI.click(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Page Header'))
WebUI.click(findTestObject('Object Repository/Equotes CONV/COVERAGE DETAILS/Continue Coverage Details'))

WebUI.delay(5)

//APPLICATION REVIEW


WebUI.click(findTestObject('Object Repository/Equotes CONV/APPLICATION REVIEW/Application Signed'))

WebUI.click(findTestObject('Object Repository/Equotes CONV/APPLICATION REVIEW/Continue Application Review'))
WebUI.delay(3)
WebUI.click(findTestObject('Object Repository/Equotes CONV/APPLICATION REVIEW/Application Reviewed'))

WebUI.click(findTestObject('Object Repository/Equotes CONV/APPLICATION REVIEW/Continue Application Review'))

WebUI.delay(5)

//PAYMENT METHOD
WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes CONV/PAYMENT METHOD/Payment Method'),
	'DIR', true)

WebUI.click(findTestObject('Object Repository/Equotes CONV/PAYMENT METHOD/Continue Payment Method'))
WebUI.delay(10)
//BENEFICIARIES DETAILS
WebUI.click(findTestObject('Object Repository/Equotes CONV/BENEFICIARIES DETAILS/Contingent No'))
WebUI.click(findTestObject('Object Repository/Equotes CONV/BENEFICIARIES DETAILS/Continue Beneficieary Details'))
WebUI.delay(5)
WebUI.click(findTestObject('Object Repository/Equotes CONV/BENEFICIARIES DETAILS/Principal No'))

WebUI.click(findTestObject('Object Repository/Equotes CONV/BENEFICIARIES DETAILS/Continue Beneficieary Details'))
WebUI.delay(5)
WebUI.click(findTestObject('Object Repository/Equotes CONV/BENEFICIARIES DETAILS/Continue Beneficieary Details'))
WebUI.delay(10)
//APPLICATION DECISION

WebUI.click(findTestObject('Object Repository/Equotes CONV/APPLICATION DECISION/Continue Application Decision'))

WebUI.delay(10)



def Sucess_Message = WebUI.getText(findTestObject('Object Repository/Equotes CONV/APPLICATION DECISION/Final Message'))
sheet.getRow(9).createCell(4).setCellValue(Sucess_Message);

if(Sucess_Message == "")
{ sheet.getRow(9).createCell(3).setCellValue("Fail");
  sheet.getRow(9).createCell(4).setCellValue("Test Case Crashed during the execution");
  sheet1.getRow(19).createCell(3).setCellValue("Fail");
  sheet1.getRow(20).createCell(3).setCellValue("Fail");
}
else {sheet.getRow(9).createCell(3).setCellValue("Pass");
	sheet1.getRow(19).createCell(3).setCellValue("Pass");
	sheet1.getRow(20).createCell(3).setCellValue("Pass");
}

	

FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))

workbook.write(outFile)

outFile.close()

WebUI.closeBrowser()
