
brw = "chrome.exe"
systemutil.CloseProcessByName brw
Kill_Process brw
WebUtil.DeleteCookies
systemutil.Run brw,"https://www.facebook.com/", , , 3
  
'rowCount = Datatable.GlobalSheet.GetRowCount

'For i = 1 To rowCount

  'Datatable.SetCurrentRow(i)
  
  un = Datatable.Value("UserName","Global")
  pw = Datatable.Value("Password","Global")

  Set homePage = HomePageInstance()
  homePage.waitForHomePageToAppear()
  homePage.setUserName(un)
  homePage.setPassword(pw)
  homePage.selectYear("1992")
  
  'Browser("creationtime:=0").Page("title:=.*").WebList("html id:=year").Select "1952"

  Set homePage = Nothing
  
'Next

  systemutil.CloseDescendentProcesses
