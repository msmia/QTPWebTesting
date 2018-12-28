
a = array(1,1,1,1,2,2,2,3,3,3,4,4,4,5,5,5)
Set oDict = CreateObject("Scripting.Dictionary")

for each ax in a
	oDict(ax) = 0
next

print join(oDict.Keys())

Set oDict = Nothing

exittest

brw = "chrome.exe"
systemutil.CloseProcessByName brw
ClearTempFolder
Kill_Process brw
Kill_Process "sublime_text.exe"
url = "www.facebook.com"
systemutil.Run brw,url, , , 3

call logger("", "Test started.")
call logger("","==================")

rowCount = Datatable.GlobalSheet.GetRowCount

For i = 1 To rowCount

  Datatable.SetCurrentRow(i)
  
  un = Datatable.Value("UserName","Global")
  pw = Datatable.Value("Password","Global")
  yr = Datatable.Value("year","Global")

  'Instantiate Home Page
  Set homePage = HomePageInstance()
  homePage.waitForHomePageToLoad()
  homePage.setUserName(un)
  homePage.setPassword(pw)
  homePage.selectYear(yr)
  
  Set homePage = Nothing
  
Next

  systemutil.CloseDescendentProcesses
