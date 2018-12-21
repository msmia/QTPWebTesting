'------------
'QTP Startups
'------------

'Kill processes before test execution starts'
Call Kill_Process("excel.exe")
Call Kill_Process("iexplore.exe")
Call Kill_Process("uft.exe")
Call Kill_Process("chrome.exe")

'driver path
projectDir = ProjectDirectory()
'environment path
environmentDir = projectDir&"\"&"environment"
'libraries path
libDir = projectDir&"\"&"libraries"
''object_repositories path
repoDir = projectDir&"\"&"object_repositories"
'test_data path
testDataDir = projectDir&"\"&"test_data" 

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

QTPpath = repoDir&"\PageObjectModel" '(((((((((((((READ TEST CASE NAMES FROM EXCEL)))))))))))))'

qtApp.Open QTPpath, True ' Open a Script to execute the Automation test cases.

'Load function libraries'
Set qtLibraries = qtApp.Test.Settings.Resources.Libraries ' Get the libraries collection object	
qtLibraries.RemoveAll
Dim functionLib1, functionLib2
functionLib1 = libDir&"\generic_functions\GenericFunctions.txt"
functionLib2 = libDir&"\ObjectLocatorClasses\HomePage.txt"
qtLibraries.Add functionLib1, 1 
qtLibraries.Add functionLib2, 1 



Set qtTest = qtApp.Test

qtTest.Run ' Run the test

qtTest.Close ' Close the test

'Now Close the QTP
qtApp.Quit ' Quit QuickTest

'Free the object holders
'Set qtResultsOpt = Nothing ' Release the Run Results Options object
Set qtTest = Nothing ' Release the Test object
Set qtApp = Nothing ' Release the Application object
'Set qtRepositories = Nothing ' Release the action's shared repositories collection



Public Function Kill_Process(strProgramName)

	'ex: notepad.exe
	'ex: AcroRd32.exe
	'ex: excel.exe

	Set WMI = GetObject("winmgmts:\\")
	Set allItem = WMI.ExecQuery("Select * from Win32_Process Where Name = "&"'"&strProgramName&"'")	
	
	For Each item in allItem
	   	item.Terminate()
	Next

End Function


'********************************************************************************************
' Function: Get the project name
' input: projectFolderName (This is the main folder under which the entire project resides)
'********************************************************************************************
Function ProjectDirectory()
	
	On error Resume Next

	scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	projectFolderName = "QTPWebTesting"
	
	If not isEmpty(projectFolderName) Then
	   If instr(1, scriptdir, projectFolderName) > -1 Then
		  demiliter = split(scriptdir, projectFolderName)
		  ProjectDirectory = demiliter(0) & projectFolderName
	   End If
	Else
		  reporter.ReportEvent micFail, "Error number: "& err.number& " Error description: " & err.description, "" 
	End If
	
	On Error Goto 0
	
End Function