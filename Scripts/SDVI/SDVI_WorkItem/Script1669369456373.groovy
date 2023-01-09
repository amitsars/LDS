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

WebUI.callTestCase(findTestCase('SDVI/Scanning Rulebook SDVI'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.delay(10)


FileInputStream file = new FileInputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheet('SDVI')
XSSFSheet sheet1 = workbook.getSheet('TC_Status')


'Read data from excel'
String Policy_Number = sheet.getRow(8).getCell(1).getStringCellValue()
String Workitemid = sheet.getRow(8).getCell(3).getStringCellValue()
String Claim_Number = sheet.getRow(9).getCell(1).getStringCellValue()
String First_Name = sheet.getRow(10).getCell(1).getStringCellValue()
String Middle_Initial = sheet.getRow(11).getCell(1).getStringCellValue()
String Last_Name = sheet.getRow(12).getCell(1).getStringCellValue()
String Gender = sheet.getRow(13).getCell(1).getStringCellValue()
String DOB = sheet.getRow(14).getCell(1).getStringCellValue()
String Country = sheet.getRow(15).getCell(1).getStringCellValue()
String Address1 = sheet.getRow(16).getCell(1).getStringCellValue()
String Address2 = sheet.getRow(17).getCell(1).getStringCellValue()
String Address3 = sheet.getRow(18).getCell(1).getStringCellValue()
String City = sheet.getRow(19).getCell(1).getStringCellValue()
String State = sheet.getRow(20).getCell(1).getStringCellValue()
String Zip_Code = sheet.getRow(21).getCell(1).getNumericCellValue()
String Phone_Number = sheet.getRow(22).getCell(1).getNumericCellValue()
String Email_Address = sheet.getRow(23).getCell(1).getStringCellValue()
//String Date_of_Death = sheet.getRow(24).getCell(1).getStringCellValue()
//String Payment_Sent_with_Application = sheet.getRow(25).getCell(1).getStringCellValue()
String SSN_Matched = sheet.getRow(26).getCell(1).getStringCellValue()
String Claim_Number_Matched = sheet.getRow(27).getCell(1).getStringCellValue()
String Name_Matched = sheet.getRow(28).getCell(1).getStringCellValue()



//WebDriver driver = new ChromeDriver()

WebUI.maximizeWindow()

WebUI.delay(5)

//WebUI.callTestCase(findTestCase('SDVI Workflow Page/Add MOD HEADER'), [:], FailureHandling.STOP_ON_FAILURE)
WebUI.navigateToUrl('http://ec04-vc02-web01:9303/underwriting')


WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Link_LISSIAflow'))
//if(Policy_Number != "") 
	//{
try {
		WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Reference_reference'),
	Policy_Number)

		WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Search Additional References_actionwo_9eb6ef'))

		WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/a_Show Details'))
		sheet1.getRow(3).createCell(3).setCellValue("Pass");
}
catch(Exception E) {
	sheet1.getRow(3).createCell(3).setCellValue("Fail");
	file.close()
	FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))
	workbook.write(outFile)
	outFile.close()
}
	//}
//else
	//{
//	WebUI.mouseOver(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Reference_reference'))
//	WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/a_advanceSearch'))
//	WebUI.delay(2)
//	WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/workitemid'),
//		Workitemid)
//	
//	WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/workitemidsearch'))
//	WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/workitemidview'))
//	
//	WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/a_Show Details'))
//	}

WebUI.delay(10)
WebUI.enableSmartWait()

try {
WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Claim Number'),
	Claim_Number)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_First Name_itemMetadata13.value'),
	First_Name)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Middle Initial_itemMetadata14.value'),
	Middle_Initial)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Last Name_itemMetadata15.value'),
	Last_Name)

//if
WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/label_Male'))

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_MMDDYYYY_itemMetadata17.value'),
	DOB)

//WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/span_NY'))

//WebUI.selectOptionByValue(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Country'),
//	Country, true)

//WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/PreselectCountry'))
//WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Country'))


WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Address1'),
	Address1)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Address2'),
	Address2)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Address3'),
	Address3)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_City_itemMetadata24.value'),
	City)

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/span_-- Please Select --'))

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/span_NY'))

WebUI.selectOptionByValue(findTestObject('Object Repository/Page_Home-LISSIAFLOW/select_-- Please Select --    AA    AE    A_2f9a67'),
	State, true)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Zip Code_itemMetadata27.value'),
	Zip_Code)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Phone Number_itemMetadata28.value'),
	Phone_Number)

WebUI.setText(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_Email Address_itemMetadata29.value'),
	Email_Address)

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/input_MMDDYYYY_itemMetadata30.value'))

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/label_No'))

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/SSN_Matched'))

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Claim_Matched'))

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/Name_Matched'))

WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/button_Open eApp'))
sheet1.getRow(4).createCell(3).setCellValue("Pass");
}
catch(Exception E) {
	sheet1.getRow(4).createCell(3).setCellValue("Fail");
	file.close()
	FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))
	workbook.write(outFile)
	outFile.close()
}

WebUI.delay(20)

//WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/a_Enter Application Information'))

if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Home-LISSIAFLOW/a_Enter Application Information'), 20))
	 {
	WebUI.click(findTestObject('Object Repository/Page_Home-LISSIAFLOW/a_Enter Application Information'))
	sheet1.getRow(5).createCell(3).setCellValue("Pass");
   } 
   else {
	sheet.getRow(8).createCell(5).setCellValue("FAILED - Enter Application URL not found");
	sheet1.getRow(5).createCell(3).setCellValue("Fail");
   }
//sheet.createRow(13).createCell(1).setCellValue('Something is wrong with SDVI WI')

file.close()

FileOutputStream outFile = new FileOutputStream(new File('C:/Users/Amit.Sarswat.EXLDEMO/Desktop/Katalon_Studio_PE_Windows_64-8.5.0/Excel/TestDataUtil.xlsx'))

workbook.write(outFile)

outFile.close()

//WebUI.closeBrowser()
