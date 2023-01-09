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

WebUI.callTestCase(findTestCase('SDVI/SDVI_WorkItem'), [:], FailureHandling.STOP_ON_FAILURE)

//WebUI.navigateToUrl('http://ec04-vc02-web01:9303/equotes/nextPage!input.action?interviewId=497555&interviewToken=9QQNYGYtAXQwdBDxQwM3')
WebUI.delay(10)

// JAI SHREE RAM

FileInputStream file = new FileInputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheet('SDVI')
XSSFSheet sheet1 = workbook.getSheet('TC_Status')


'Read data from excel'

String Insurance_Amount = sheet.getRow(30).getCell(1).getStringCellValue()
String Plan_of_Insurance = sheet.getRow(31).getCell(1).getStringCellValue()
String Effective_Date = sheet.getRow(32).getCell(1).getStringCellValue()
String Date_of_Discharge = sheet.getRow(34).getCell(1).getStringCellValue()
String Date_of_Recent_Letter = sheet.getRow(34).getCell(3).getStringCellValue()
String Payment_Method = sheet.getRow(38).getCell(1).getStringCellValue()
String Amount_Received = sheet.getRow(42).getCell(1).getStringCellValue()
String PB1_FirstName = sheet.getRow(45).getCell(1).getStringCellValue()
String PB1_MiddleName = sheet.getRow(46).getCell(1).getStringCellValue()
String PB1_LastName = sheet.getRow(47).getCell(1).getStringCellValue()
String PB1_Relationship = sheet.getRow(48).getCell(1).getStringCellValue()
String PB1_AddressLine1 = sheet.getRow(49).getCell(1).getStringCellValue()
String PB1_AddressLine2 = sheet.getRow(50).getCell(1).getStringCellValue()
String PB1_AddressLine3 = sheet.getRow(51).getCell(1).getStringCellValue()
String PB1_City = sheet.getRow(52).getCell(1).getStringCellValue()
String PB1_ZipCode = sheet.getRow(53).getCell(1).getStringCellValue()
String PB1_Email = sheet.getRow(54).getCell(1).getStringCellValue()
String PB1_State = sheet.getRow(55).getCell(1).getStringCellValue()
String CB1_FirstName = sheet.getRow(57).getCell(1).getStringCellValue()
String CB1_MiddleName = sheet.getRow(58).getCell(1).getStringCellValue()
String CB1_LastName = sheet.getRow(59).getCell(1).getStringCellValue()
String CB1_Relationship = sheet.getRow(60).getCell(1).getStringCellValue()
String CB1_AddressLine1 = sheet.getRow(61).getCell(1).getStringCellValue()
String CB1_AddressLine2 = sheet.getRow(62).getCell(1).getStringCellValue()
String CB1_AddressLine3 = sheet.getRow(63).getCell(1).getStringCellValue()
String CB1_City = sheet.getRow(64).getCell(1).getStringCellValue()
String CB1_ZipCode = sheet.getRow(65).getCell(1).getStringCellValue()
String CB1_Email = sheet.getRow(66).getCell(1).getStringCellValue()
String CB1_State = sheet.getRow(67).getCell(1).getStringCellValue()

WebUI.switchToWindowTitle('eApp')

//APPLICATION INFORMATION
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/APPLICATION INFORMATION/Application Decision Continue button'))

//COVERAGE DETAILS
WebUI.enableSmartWait()
WebUI.setText(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Amount of Insurance'),
	Insurance_Amount )

WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Insurance Plan'),
	Plan_of_Insurance, true)

//WebUI.click(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Effective Date'))


WebUI.setText(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Effective Date'),
	Effective_Date )
WebUI.click(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Page Header'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Application Information Reviewed Yes'))

WebUI.enableSmartWait()

WebUI.click(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Continue button'))

WebUI.enableSmartWait()

WebUI.click(findTestObject('Object Repository/Equotes SDVI/COVERAGE DETAILS/Continue button'))

WebUI.delay(5)
//ELIGIBILITY DETERMINATION
WebUI.enableSmartWait()

//WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Date of Discharge'))
WebUI.setText(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Date of Discharge'),
	Date_of_Discharge)

WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Page Header'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Character Discharge Yes'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Continue button'))

WebUI.delay(3)

WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/ServiceConnectedDisabilityGranted Yes'))


WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Continue button'))


WebUI.delay(3)
WebUI.enableSmartWait()

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Date Most Recent Award Letter'),
	Date_of_Recent_Letter)

WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Page Header'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Continue button'))

WebUI.delay(3)
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/ELIGIBILITY DETERMINATION/Continue button'))

WebUI.delay(10)
WebUI.enableSmartWait()
//MEDICAL DETERMINATION
WebUI.click(findTestObject('Object Repository/Equotes SDVI/MEDICAL DETERMINATION/NSC Conditions No'))
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/MEDICAL DETERMINATION/Continue button'))

WebUI.delay(5)

//PAYMENT METHOD

WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/Payment Method'),
	Payment_Method, true)

WebUI.click(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/Requested Waiver No'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/Cash Received with Application Yes'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/Continue button'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/SufficientFundsAvailable Yes'))

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/Amount Received'),
	Amount_Received )

WebUI.click(findTestObject('Object Repository/Equotes SDVI/PAYMENT METHOD/Continue button'))

WebUI.delay(3)

//BENEFICIARIES DETAILS
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/Principal Divide Share Equally Yes'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/Add Contingent Beneficiary Yes'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/Continue button'))

//PB1 details
WebUI.enableSmartWait()
WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 FirstName'),
	PB1_FirstName )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Middle Name'),
	PB1_MiddleName )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Last Name'),
	PB1_LastName )

WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Relation Ship'),
	PB1_Relationship , true)

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Address Line 1'),
	PB1_AddressLine1)

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Address Line 2'),
	PB1_AddressLine2)

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Address Line 3'),
	PB1_AddressLine3 )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 City'),
	PB1_City )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Zip Code'),
	PB1_ZipCode )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 Email'),
	PB1_Email )

WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB Divide Share Equally Yes'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/Continue button'))
WebUI.enableSmartWait()
WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/PB1 State'),
	PB1_State, true)

//CB1 details

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 First Name'),
	CB1_FirstName )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Middle Name'),
	CB1_MiddleName )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Last Name'),
	CB1_LastName )

WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Relationship'),
	CB1_Relationship , true)

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Address Line 1'),
	CB1_AddressLine1 )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Address Line 2'),
	CB1_AddressLine2)

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Address Line 3'),
	CB1_AddressLine3)

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 City'),
	CB1_City )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Zip Code'),
	CB1_ZipCode )

WebUI.setText(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 Email'),
	CB1_Email)

WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/Continue button'))
WebUI.enableSmartWait()
WebUI.selectOptionByValue(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/CB1 State'),
	CB1_State , true)

WebUI.click(findTestObject('Object Repository/Equotes SDVI/BENEFICIARIES DETAILS/Continue button'))

//WAIVER DETERMINATION
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/WAIVER DETERMINATION/Sufficient Evidence To Process Waiver No'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/WAIVER DETERMINATION/Continue button'))
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/WAIVER DETERMINATION/Inviting Veterian for Waiver No'))

WebUI.click(findTestObject('Object Repository/Equotes SDVI/WAIVER DETERMINATION/Continue button'))
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/WAIVER DETERMINATION/Continue button'))

//APPLICATION DECISION
WebUI.enableSmartWait()
WebUI.click(findTestObject('Object Repository/Equotes SDVI/APPLICATION DECISION/Continue button'))

WebUI.delay(20)



def Sucess_Message = WebUI.getText(findTestObject('Object Repository/Equotes SDVI/APPLICATION DECISION/Final Message'))
sheet.getRow(9).createCell(4).setCellValue(Sucess_Message);


if(Sucess_Message == "")
{ sheet.getRow(9).createCell(3).setCellValue("Fail");
  sheet.getRow(9).createCell(4).setCellValue("Test Case Crashed during the execution");\
  sheet1.getRow(6).createCell(3).setCellValue("Fail");
  sheet1.getRow(7).createCell(3).setCellValue("Fail");
}
else {sheet.getRow(9).createCell(3).setCellValue("Pass");
	sheet1.getRow(7).createCell(3).setCellValue("Pass");
	sheet1.getRow(6).createCell(3).setCellValue("Pass");
}

	

FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))

workbook.write(outFile)

outFile.close()


WebUI.closeBrowser()


