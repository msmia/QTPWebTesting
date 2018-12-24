
brw = "chrome.exe"
systemutil.CloseProcessByName brw
Kill_Process brw
WebUtil.DeleteCookies
systemutil.Run brw,"https://www.facebook.com/", , , 3
  
rowCount = Datatable.GlobalSheet.GetRowCount

For i = 1 To rowCount

  Datatable.SetCurrentRow(i)
  
  un = Datatable.Value("UserName","Global")
  pw = Datatable.Value("Password","Global")
  yr = Datatable.Value("year","Global")

  Set signupPage = signupPageInstance()
  signupPage.waitForSignupPageToLoad()
  signupPage.setUserName(un)
  signupPage.setPassword(pw)
  signupPage.selectYear(yr)
  

  Set signupPage = Nothing
  
Next

  systemutil.CloseDescendentProcesses
