'****************'
'* Driver Script '
'****************'

'Kill processes'
Call Kill_Process("excel.exe")
Call Kill_Process("iexplore.exe")
Call Kill_Process("uft.exe")
Call Kill_Process("chrome.exe")

'============================================================'
iResultSheetRowCounter = 2
iTestCaseExecuting = 1
iTotalPassed = 0
iTotalFailed = 0
iTotalOthers = 0




'============================================================'

'driver path
projectDir = ProjectDirectory()
'environment path
environmentDir = projectDir&"\environment"
'Execution logs'
logFileDir = projectDir&"\execution_logs\"
'libraries path
libDir = projectDir&"\libraries"
'page_objects path'
pageObjectsDir = projectDir&"\page_objects"
'Detail Results
detailedResultsDir = projectDir&"\results\detailed_qtp_results\"
'Summarizied Results'
summarizedResultsDir = projectDir&"\results\summarized_results\"
''page_objects path
scriptsDir = projectDir&"\scripts"
'test_data path
testDataDir = projectDir&"\test_data" 

'Create a QTP Object
Dim qtApp, qtTest

Set qtApp = CreateObject("QuickTest.Application") 'Create the Application object

'Launch qtp and make it visible'
qtApp.Launch
qtApp.Visible = True

' Set QuickTest run options
'qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
'qtApp.Options.Run.RunMode = "Fast"
qtApp.Options.Run.ViewResults = False

'************************************************'
'Test case names in an array'					 '
'Loop through each testcase name and run the test'
'************************************************'
Dim testCaseNamesArray

testCaseNamesArray = array("SigninTest")
For aa = 0 to ubound(testCaseNamesArray)
	qtApp.Open scriptsDir&"\"&testCaseNamesArray(aa), true
	
	
'****************************************'
'	     Function And Page Library       '
'****************************************'
	Dim qtLibraries, functionLibraries, pageObjects
	
	Set qtLibraries = qtApp.Test.Settings.Resources.Libraries
	qtLibraries.RemoveAll
	
	'Associate function libraries
	Set functionLibraries = listOfFiles(libDir)
	for i = 0 to functionLibraries.count-1
		qtLibraries.Add libDir &"\"& functionLibraries(i), 1
	next
	
	'Associate pages
	Set pageObjects = listOfFiles(pageObjectsDir)
	for i = 0 to pageObjects.count-1
		qtLibraries.Add pageObjectsDir &"\"& pageObjects(i), 1
	next
	
	'Execute test'
	Set qtTest = qtApp.Test
	qtTest.Run ' Run the test
	
	'Close the test'
	qtTest.Close 
	
Next


'Now Close the QTP
qtApp.Quit 

'Free the object holders
Set qtTest = Nothing
Set qtLibraries = Nothing
Set functionLibraries = Nothing
set pageObjects = Nothing
Set qtApp = Nothing
'Set qtResultsOpt = Nothing
'Set qtRepositories = Nothing



'*****************************************************'
'			FUNCTIONS BELOW THIS LINE	  			  '
'*****************************************************'

Public Function Kill_Process(strProgramName)
	
'ex: notepad.exe / AcroRd32.exe / excel.exe
	On error resume next
	
	Set WMI = GetObject("winmgmts:\\")
	Set allItem = WMI.ExecQuery("Select * from Win32_Process Where Name = "&"'"&strProgramName&"'")	
	For Each item in allItem
			item.Terminate()
	Next
	
	Set WMI = Nothing
	set allItem = Nothing
	On error goto 0
End Function


'*******************************'
' Function: Get the project name'
'*******************************'

Function ProjectDirectory()
	
	On error Resume Next
	
	scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	projectFolderName = "QTPWebTesting"
	
	If not isEmpty(projectFolderName) Then
		If instr(1, scriptdir, projectFolderName) > - 1 Then
			demiliter = split(scriptdir, projectFolderName)
			ProjectDirectory = demiliter(0) & projectFolderName
		End If
	Else
		reporter.ReportEvent micFail, "Error number: "& err.number& " Error description: " & err.description, "" 
	End If
	
	On Error Goto 0
End Function


'****************************************'
'Function: listOfFiles					 '
'input: It takes a folder path as input  '
'output: arraylist of files in the folder'
'****************************************'

Function listOfFiles(strFolderDirectory)
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	set libFolder = fso.GetFolder(strFolderDirectory)
	
	set filesArrayList = CreateObject("System.Collections.ArrayList")
	
	For each file in libFolder.files
		filesArrayList.add(file.name)
	Next
	
	Set listOfFiles = filesArrayList
	
	Set filesArrayList = Nothing
	Set libFolder = Nothing
	Set fso = Nothing
	
End Function
