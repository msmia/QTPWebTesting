


'Start loggin information
call logger("", "============")
call logger("", "Test started.")
call logger("", "============")

'Declare variables
brw = "chrome.exe"


'Close appications
ClearTempFolder
Call Kill_Process("excel.exe")
Call Kill_Process("iexplore.exe")
Call Kill_Process("chrome.exe")
Call Kill_Process("sublime_text.exe")

url = "www.facebook.com"
systemutil.Run brw,url, , , 3


'Prepare the test data
projectFolder = "C:\QA\GIT\QTPWebTesting\"
testDataFolderPath = projectFolder & "test_data\"

xlTestCaseFile = "TestCaseNames.xlsx"
xlTestDataFile = "TestData.xlsx"

'Import test case names
shTCnames = "TC_ID"
datatable.AddSheet shTCnames
datatable.ImportSheet testDataFolderPath & xlTestCaseFile, shTCnames, shTCnames


'Instantiate required classes
Set homePage = HomePageInstance()



'Test Cases
'==========
'==========
rowCount = Datatable.GetSheet(shTCnames).GetRowCount

For mainLoop = 1 To rowCount

  Datatable.GetSheet(shTCnames).SetCurrentRow(mainLoop)
  
  tcName = datatable.Value("TC_ID", shTCnames)
  execFlag = datatable.Value("ExecutinFlag", shTCnames)
  
  If Ucase(execFlag) = "Y" Then
  
  	If Trim(tcName) = "001_GoodSignin" Then
  	
  	   Call logger("","001_GoodSignin is not ready yet.")
  	   
  	ElseIf Trim(tcName) = "002_BadSignin" Then
  	
  	   shBadSignin = "BadSignin"
	   datatable.AddSheet shBadSignin
	   datatable.ImportSheet testDataFolderPath & xlTestDataFile, shBadSignin, shBadSignin
  	   Call badSignin(homePage, shBadSignin)
  	   
  	End If
  End If
  
Next

'Release classes memories
Set homePage = Nothing

