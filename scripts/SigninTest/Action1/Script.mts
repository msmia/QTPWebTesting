'url = "www.facebook.com"
'systemutil.Run "chrome.exe",url, , , 3
'exittest
	


'***********'
'*	Setup  *'
'***********'


	'Close appications
	Call Kill_Process("excel.exe")
	Call Kill_Process("iexplore.exe")
	Call Kill_Process("chrome.exe")
	Call Kill_Process("sublime_text.exe")


	'Project path
	projectPath = ProjectDirectory("QTPWebTesting")
	'test data path
	testDataFolderPath = projectPath & "test_data\"

	
	Wait (2)


	'Open Browser
	systemutil.Run Environment.Value("chrome"), Environment.Value("url"), , , 3
	
	
	'Instantiate classes
	Set homePage = HomePageInstance()


'**********'
'*	Test  *'
'**********'
	
	
	'Start loggin information
	Call logger("", "============ Test started. ============")
	
		
	'Import test case names sheet
	shTCnames = "TestCaseNames"
	ExcelFile = testDataFolderPath & "Config.xlsx"
	fnImportSheet ExcelFile, shTCnames
	
	rowCount = Datatable.GetSheet(shTCnames).GetRowCount
	
	For mainLoop = 1 To rowCount
	
	  Datatable.GetSheet(shTCnames).SetCurrentRow(mainLoop)
	  
	  tcName   = datatable.Value("TC_ID", shTCnames)
	  execFlag = datatable.Value("ExecutinFlag", shTCnames)
	  
	 'Flag = Y
	  If Ucase(execFlag) = "Y" Then
	  
	  	If Trim(tcName) = "001_GoodSignin" Then
	  	
	  	   'Script not ready yet
	  	 	
	  	ElseIf Trim(tcName) = "002_BadSignin" Then
	  	
		   shBadSignin = "BadSignin"
		   ExcelFile = testDataFolderPath & "TestData.xlsx"
		   fnImportSheet ExcelFile, shBadSignin
	  	   Call badSignin(homePage, shBadSignin)

	  	End If
	  	
	  End If
	  
	Next
	
	
	'Delete imported sheets
	Datatable.DeleteSheet shBadSignin
	Datatable.DeleteSheet shTCnames
	
	'Release classes memories
	Set homePage = Nothing

