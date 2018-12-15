
'-----------------------------------------------------------
'Script Description        : Driver class for the framework
'Test Tool/Version         : HP Quick Test Professional 12+
'Test Tool Settings        : N.A.
'Application Automated     : N.A.
'Author                    : Sharif Mia
'Date Created              : Oct 2017
'Date Modified             : Oct 2017
'-----------------------------------------------------------

iResultSheetRowCounter = 2 : iTestCaseExecuting = 1 : iTotalPassed = 0 : iTotalFailed = 0 : iTotalOthers = 0


'sFrameworkFolder = "[ALM] Subject\Automation\Internal Apps\"
'sTestCaseFolder = sFrameworkFolder & "Dealer Rating\"

''######################################################
sFrameworkFolder = "H:\quick_test_result_practice_UFT\"
sTestCaseFolder = sFrameworkFolder & "TestCases\"
sQTPResultsPathOrig = sFrameworkFolder & "Results\DetailedQTPResults\"
sBatchRunPath = sFrameworkFolder & "Results\SummarizedResults\"
sBatchSheetPath = sFrameworkFolder & "TestCaseNames.xlsx"
sBatchSheetName = "TC_ID"
sResultSheetName = "Sheet1"

'-----------*****----------
'		Call Functions	  '
'-----------*****----------

'Kill processes before test execution starts'
Call Kill_Process("excel.exe")
Call Kill_Process("iexplore.exe")
Call Kill_Process("uft.exe")

'Create time stamp
call fnTimeStamp(sTimeStamp)

'Create the result excel sheet
Call fnCreateResultExcelSheet(sBatchRunPath, sTimeStamp)

'Execute the testCases
Call fnExecuteTestCases(sBatchSheetPath, sBatchSheetName)


'Kill processes after test execution ends'
Call Kill_Process("excel.exe")
Call Kill_Process("iexplore.exe")
Call Kill_Process("uft.exe")

'Test is complete'
msgbox "The test is completed."






'---------*****------
'		Functions	'
'---------*****------



''######################################################
Function fnExecuteTestCases(sBatchSheetPath, sBatchSheetName)

Dim qtpApp, qtTest

'Create qtp object'
Set qtpApp = CreateObject("QuickTest.Application") 

'Now Launch the QTP
If  qtpApp.launched <> True then
	Wscript.SLeep 3000 : qtpApp.Launch : Wscript.SLeep 3000
End If

'Make qtp visible
qtpApp.Visible =True
Wscript.SLeep 1000

' Set QuickTest run options
qtpApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtpApp.Options.Run.RunMode = "Normal"
qtpApp.Options.Run.ViewResults = False


'Read test name data from excel based on execute flag
Set xl_Batch = CreateObject("Excel.Application")
	xl_Batch.WorkBooks.Open(sBatchSheetPath)
Set xlSheet = xl_Batch.ActiveWorkbook.Worksheets(sBatchSheetName)

'Get the row count
Row = xlSheet.UsedRange.Rows.Count

For iR = 2 to Row	
	
	'Execution flag
	executionFlag = xl_Batch.Cells(iR, 2).Value

	If Ucase(Trim(executionFlag)) = "Y" Then

		'Run the TC and Update Results
		
		iTestCaseExecuting = iTestCaseExecuting + 1
		
		'Get testcase Name
		sTestCaseName = xl_Batch.Cells(iR, 1).Value	

		'Get QTP script path
		strScriptPath = sTestCaseFolder & sTestCaseName
		
		'Open the Test Case in Read-Only mode
		qtpApp.Open strScriptPath, True
		WScript.Sleep 2000

		'Associate function Libraries'
		Set objLib = qtpApp.Test.Settings.Resources.Libraries
		If objLib.Find("H:\cna_website_automatin\functions for cna projects\All CNA custom functions by Sharif.qfl.txt") = -1 Then 
  			objLib.Add "H:\cna_website_automatin\functions for cna projects\All CNA custom functions by Sharif.qfl.txt", 1 
		End If

		If objLib.Find("H:\cna_website_automatin\functions for cna projects\All custom functions by Sharif.qfl.txt") = -1 Then 
  			objLib.Add "H:\cna_website_automatin\functions for cna projects\All custom functions by Sharif.qfl.txt", 1 
		End If		

		'set run settings for the test
		Set qtpTest = qtpApp.Test
		
		'Instruct QuickTest to perform next step when error occurs
		qtpTest.Settings.Run.OnError = "NextStep"

		'Create the Run Results Options object
		Set qtpResult = CreateObject("QuickTest.RunResultsOptions")

		'Set the results location
		sQTPResultsPath = sQTPResultsPathOrig & sTestCaseName  & "_" & sTimeStamp
		qtpResult.ResultsLocation = sQTPResultsPath

		'Find start date time before test execution
		exeDate = Date
		exeStartTime = Time()
		exeStartTimer = Timer()

		'Run the test
		WScript.Sleep 2000
		qtpTest.Run qtpResult
		
		'Find end date time after test execution
		exeEndTime = Time()
		exeEndTimer = Timer()

		'Divide the second to minutes'
		exeDuration = exeEndTimer - exeStartTimer	
		If 	exeDuration >= 60 then
			exeDuration = exeDuration / 60
			exeDuration = round(exeDuration,2)
		End If	 

		'Find Run Status
		sRunStatus = qtpTest.LastRunResults.Status
		Select Case sRunStatus
			Case "Passed"   		
				iTotalPassed = iTotalPassed + 1
			Case "Failed"			   
				iTotalFailed = iTotalFailed + 1
			Case Else				 	
				iTotalOthers = iTotalOthers + 1
		End Select

		'Call function to Save the result
		Call fnSaveTestCaseResult(sResultSheetName, sBatchRunPath, sTimeStamp, iResultSheetRowCounter, sTestCaseName, executionFlag, sRunStatus, sQTPResultsPath, exeDate, exeStartTime, exeEndTime, exeDuration)		
		'+++++++++++++++++++++++++++++++++
		
		ElseIf Ucase(Trim(executionFlag)) = "N" Then
			'Get TC Name
			sTestCaseName = xl_Batch.Cells(iR, 1).Value
			'Call function to Save the result
			Call fnSaveTestCaseResult(sResultSheetName, sBatchRunPath, sTimeStamp, iResultSheetRowCounter, sTestCaseName, executionFlag, sRunStatus, sQTPResultsPath, exeDate, exeStartTime, exeEndTime, exeDuration)			

		ElseIf xl_Batch.Cells(iR, 2).Value = "" Then
			Exit For		
	End If

		
		'Kill process'
		Call Kill_Process("iexplore.exe")

		
		'Display auto message of the testing pass / fail status'
		Call autoCloseMsgbox(iTotalPassed, iTotalFailed, iTotalOthers)
		'+++++++++++++++++++++++++++++++++		
		
		Wscript.SLeep 4000
	
Next 

	'Delete references
	xl_Batch.Quit()
	qtpApp.Quit()
	Set wb_Batch = Nothing
	Set xlSheet = Nothing
	Set xl_Batch = Nothing
	Set qtpApp = Nothing

End Function

''######################################################

''######################################################

'Create time stamp'
Function fnTimeStamp(sTimeStamp)
	sTime = Now()
	sTime = replace(sTime, "/","_")
	sTime = replace(sTime, " ","_")
	sTime = replace(sTime, ":","_")
	sTimeStamp = sTime
End Function


''######################################################

''######################################################

'Create excel result sheet'
Function fnCreateResultExcelSheet(sBatchRunPath, sTimeStamp)

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = False
	objExcel.DisplayAlerts = False
	Set objWorkbook = objExcel.Workbooks.Add()

	objExcel.Columns("A:A").ColumnWidth = 41
	objExcel.Columns("B:B").ColumnWidth = 10
	objExcel.Columns("C:C").ColumnWidth = 40
	objExcel.Columns("D:D").ColumnWidth = 12
	objExcel.Columns("E:E").ColumnWidth = 12
	objExcel.Columns("F:F").ColumnWidth = 12
	objExcel.Columns("G:G").ColumnWidth = 12

	'Set Calibry Font for the excel sheet
	Set objRange = objExcel.Range("A1:L100")
	objRange.Font.Size = 10
	objRange.Font.Name = "Calibri"
	objRange.Font.Bold = FALSE
	Set objRange = Nothing

	Set objRange = objExcel.Range("A1:H1")
	objRange.Font.Size = 10
	objRange.Font.Bold = TRUE
	Set objRange = Nothing

	'Set Header
	objExcel.Cells(1, 1).Value = "TestCase_Name"
	objExcel.Cells(1, 2).Value = "Status"
	objExcel.Cells(1, 3).Value = "Test Results Path"
	objExcel.Cells(1, 4).Value = "Execution Date"
	objExcel.Cells(1, 5).Value = "Start Time"
	objExcel.Cells(1, 6).Value = "End Time"
	objExcel.Cells(1, 7).Value = "Duration"

	'Save and close excel
	objWorkbook.SaveAs sBatchRunPath & sTimeStamp & ".xlsx"
	objExcel.Quit
	Set objWorkbook = Nothing
	Set objExcel = Nothing

End Function


''######################################################

''######################################################


'Save test execution result to excel'
Function fnSaveTestCaseResult(sResultSheetName, sBatchRunPath, sTimeStamp, iResultSheetRowCounter, sTestCaseName, executionFlag, sRunStatus, sQTPResultsPath, exeDate, exeStartTime, exeEndTime, exeDuration)


	'If the test case execution flag was "Y"
	If Ucase(Trim(executionFlag)) = "Y" Then	
		'Open Result Sheet and update the result
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		Set objWorkbook = objExcel.WorkBooks.Open(sBatchRunPath & sTimeStamp & ".xlsx")
		Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets(sResultSheetName)
		objExcel.DisplayAlerts = False
		
		'Set the results
		objExcel.Cells(iResultSheetRowCounter, 1).Value = sTestCaseName
		objExcel.Cells(iResultSheetRowCounter, 2).Font.Bold = TRUE
		objExcel.Cells(iResultSheetRowCounter, 2).Value = sRunStatus

		'passCounter = 0 : failCounter = 0 : NACounter = 0
		'Select Case 

		'Color status
		Select Case sRunStatus
			Case "NA"
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(139, 137, 137)
			Case "Passed"
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(0, 100, 0)
			Case "Failed"
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(245, 0, 0)
			Case Else
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(255, 255, 0)
		End Select
		
		'Write the values to excel result sheet'
		objExcel.Cells(iResultSheetRowCounter, 3).Value = sQTPResultsPath
		objExcel.Cells(iResultSheetRowCounter, 4).Value = exeDate
		objExcel.Cells(iResultSheetRowCounter, 5).Value = exeStartTime
		objExcel.Cells(iResultSheetRowCounter, 6).Value = exeEndTime
		objExcel.Cells(iResultSheetRowCounter, 7).Value = exeDuration
		
		'Autofit cells
		objExcel.Application.Sheets(1).Columns("A:I").AutoFit

		'Middle Alignment
		objWorkSheet.Cells(iResultSheetRowCounter, 2).HorizontalAlignment  = -4108
		objWorkSheet.Cells(iResultSheetRowCounter, 4).HorizontalAlignment  = -4108
		objWorkSheet.Cells(iResultSheetRowCounter, 5).HorizontalAlignment  = -4108
		objWorkSheet.Cells(iResultSheetRowCounter, 6).HorizontalAlignment  = -4108
		objWorkSheet.Cells(iResultSheetRowCounter, 7).HorizontalAlignment  = -4108


		
		iResultSheetRowCounter = iResultSheetRowCounter + 1
	
		'Save and close excel
		objWorkbook.Save
		objWorkbook.close
		objExcel.Quit
		Set objWorkbook = Nothing
		Set objExcel = Nothing

	ElseIf Ucase(Trim(executionFlag)) = "N" Then

		'Open Result Sheet and update the result
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		Set objWorkbook = objExcel.WorkBooks.Open(sBatchRunPath & sTimeStamp & ".xlsx")
		objExcel.DisplayAlerts = False
		
		'Set the results
		objExcel.Cells(iResultSheetRowCounter, 1).Value = sTestCaseName
		objExcel.Cells(iResultSheetRowCounter, 2).Font.Bold = TRUE
		objExcel.Cells(iResultSheetRowCounter, 2).Value = "N/A"
		
		'Color status
		Select Case sRunStatus
			Case "NA"
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(139, 137, 137)
			Case "Passed"
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(0, 100, 0)
			Case "Failed"
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(245, 0, 0)
			Case Else
				objExcel.Cells(iResultSheetRowCounter, 2).Font.Color = RGB(255, 255, 0)
		End Select
		
		objExcel.Cells(iResultSheetRowCounter, 3).Value = sQTPResultsPath
		objExcel.Cells(iResultSheetRowCounter, 4).Value = ""
		objExcel.Cells(iResultSheetRowCounter, 5).Value = ""
		objExcel.Cells(iResultSheetRowCounter, 6).Value = ""
		objExcel.Cells(iResultSheetRowCounter, 7).Value = ""
		
		iResultSheetRowCounter = iResultSheetRowCounter + 1
	
		'Save and close excel
		objWorkbook.Save
		objWorkbook.close
		objExcel.Quit
		Set objWorkbook = Nothing
		Set objExcel = Nothing		
	End If


End Function


''######################################################

''######################################################


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


''######################################################

''######################################################


Function autoCloseMsgbox(iTotalPassed, iTotalFailed, iTotalOthers)
	 
  Set objWS = createobject("wscript.shell")
	
	msgbox_message = "Total Pass - " &  iTotalPassed & vbcrlf &_
					 "Total Fail - " &  iTotalFailed & vbcrlf &_
					 "Total Others - " &  iTotalOthers

	msgbox_title =   "Auto close message box."

	timeOut = 2
	
	objWS.popup msgbox_message, timeOut, msgbox_title

  Set objWS = Nothing
    
End Function 