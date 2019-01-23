

'***********'
'*	Setup  *'
'***********'
	
	'Name of the project
	projectName = Environment.Value("projectName")
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
	'Browser
	strBrowser = Environment.Value("chrome")
	'Url
	strUrl = Environment.Value("url")
	
	'Call associateFiles(libPath)
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






''*********************'
''Get project directory'
''*********************'
'Function ProjectDirectory(projectName)
'	
'	On error Resume Next
'	
'	testActionPath = Environment.Value("TestDir")
'	'projectName = "QTPWebTesting"
'	
'	If not isEmpty(projectName) Then
'		If instr(1, testActionPath, projectName) > - 1 Then
'			demiliter = split(testActionPath, projectName)
'			ProjectDirectory = demiliter(0) & projectName & "\"
'		End If
'	Else
'		reporter.ReportEvent micFail, "Error number: "& err.number& " Error description: " & err.description, "" 
'	End If
'	
'	On Error Goto 0
'End Function
'
'
''***************'
''Associate Files'
''***************'
'Function associateFiles(folderPath)
'
'	Set fso= CreateObject("Scripting.FileSystemObject")
'	Set f = fso.GetFolder(folderPath)
'	Set fc = f.files
'	  For Each singlefile in fc
'		  strName = split(lcase(singlefile.name), ".")
'		  If instr(1, strName(1), "qfl") > - 1 OR instr(1, strName(1), "txt") > - 1 Then
'		  	file = folderPath&"\"&singlefile.name
'		  	ExecuteFile file
'		  End If
'	 Next
'	 
'	Set fso = Nothing
'	Set f = Nothing
'	Set fc = Nothing
'End Function 
