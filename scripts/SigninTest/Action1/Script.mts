

'Note: 
'Make sure the project name is = "QTPWebTesting"
'Make sure to dissociate and reassociate all the pages and function libraries.



'***********'
'*	Setup  *'
'***********'
	
	'Name of the project
	projectName = Environment.Value("projectName")
	'Browser
	strBrowser = Environment.Value("ie")
	'Url
	strUrl = Environment.Value("url")
	
	'Project path
	projectPath = ProjectDirectory(projectName)
	'environment path
	environmentPath = projectPath&"environment"
	'Execution logs'
	logFilePath = projectPath&"execution_logs"
	'libraries path
	libPath = projectPath&"libraries"
	'page_objects path'
	pageObjectsPath = projectPath&"page_objects"
	'Detail Results
	detailedResultsPath = projectPath&"results\detailed_qtp_results"
	'Summarizied Results'
	summarizedResultsPath = projectPath&"results\summarized_results"
	'page_objects path
	scriptsPath = projectPath&"scripts"
	'test data path
	testDataFolderPath = projectPath & "test_data\"
	
	'Close appications
	Call Kill_Process("excel.exe")
	Call Kill_Process("iexplore.exe")
	Call Kill_Process("chrome.exe")
	Call Kill_Process("sublime_text.exe")
	
	Wait (2)

	'Open Browser
	systemutil.Run strBrowser, strUrl, , , 3
	
	
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
	  execFlag = datatable.Value("ExecutionFlag", shTCnames)
	  
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
