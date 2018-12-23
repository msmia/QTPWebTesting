

brw = "chrome.exe"
systemutil.CloseProcessByName brw
Kill_Process brw
ClearBrowserHistory
url = "www.facebook.com"
systemutil.Run brw,url, , , 3

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
  
Next

  Set homePage = Nothing
  
  systemutil.CloseDescendentProcesses
