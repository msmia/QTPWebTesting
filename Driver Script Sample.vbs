'--------------------------------------------------------------------------------------------------
'QTP Startups
'--------------------------------------------------------------------------------------------------

'Create a QTP Object

Dim qtApp
Dim qtTest

Set qtApp = CreateObject("QuickTest.Application") 'Create the Application object

'Now Launch thye QTP, allow it to be visible
qtApp.Launch ' Launch QuickTest
qtApp.Visible = True ' Set QuickTest to be visible

' Set QuickTest run options
qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtApp.Options.Run.RunMode = "Fast"
qtApp.Options.Run.ViewResults = TRUE

QTPpath = "C:\QA\GIT\QTPWebTesting\object_repositories\BadUser"

qtApp.Open QTPpath, True ' Open a Script to execute the Automation test cases.

'Set qtLibraries = qtApp.Test.Settings.Resources.Libraries ' Get the libraries collection object	

'qtLibraries.Add "D:\GP_Raju_QTP\QTP Framework\Libraries\Generic_Funcs\Login.vbs", 1 ' Add the library to the collection


Set qtTest = qtApp.Test

qtTest.Run ' Run the test

qtTest.Close ' Close the test

'Now Close the QTP
qtApp.Quit ' Quit QuickTest

'Free the object holders
Set qtResultsOpt = Nothing ' Release the Run Results Options object
Set qtTest = Nothing ' Release the Test object
Set qtApp = Nothing ' Release the Application object
Set qtRepositories = Nothing ' Release the action's shared repositories collection
Set qtApp = Nothing ' Release the Application object
